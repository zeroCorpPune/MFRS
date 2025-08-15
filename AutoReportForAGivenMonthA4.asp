<!--#include file="mCommon.asp"-->

<%
'response.write Request.queryString("txtSQL")
'response.write "Wip"
'response.end

'	If trim(Session("User")) <> "RCSVIP"  Then
'			Response.write "<Font Face=Verdana Size=-1 Color=#FF0000><B>Work in progress. Please try later.</B></Font><HR>"
'			Response.End
'	End if

	If trim(Session("User")) = "" Or trim(Session("PINId")) = "" Then
			Response.write "<Font Face=Verdana Size=-1 Color=#FF0000><B>Session Time Out Re-login.</B></Font><HR>"
			Response.End
	End if

	If Session("USERTYPE") <> "SU" Then
			Response.write "<Font Face=Verdana Size=-1 Color=#FF0000><B>Invalid operation.</B></Font><HR>"
			Response.End
	End if


call setConnection
Call EnterInLog(Session("User"),Trim(Session("Company_Code")),"R","Group Result Summary Report.")
	'
sSql=""
PageNumber=0
DispCondRpt=""
CompID=""
CompanyName=""
Remarks=""
cMM=""
cMM1=""
cYY=""
ssSql=""
AnnexsSql=""
SectorNo=""
TheReportYear=Request.QueryString("theyear")

If CInt(Request.QueryString("themonth"))<=9 Then
	TheReportMonth="0" & CInt(Request.QueryString("themonth"))
Else 
	TheReportMonth=Request.QueryString("themonth")
End If 

MonthsArr=Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec")
FullMonthsArr=Array("January","February","March","April","May","June","July","August","September","October","November","December")


'response.write MonthsArr(1)
'response.write cint(TheReportMonth)
'response.write TheReportYear
'response.end

%>
<Html>



	<Title>
	Group Result Summary Report
	</Title>
<STYLE>



/*--For Page--*/
#footer {
    clear: both;
    position: relative;
    height: 40px;
    margin-top: -40px;
}
/*--*/

.no-top-border_thick {
    border-right: solid 1px black !important;
	border-left: solid 3px black !important;
	border-top: solid 0px black !important;
	border-bottom: solid 1px black !important;
	
	
}

.no-top-right-border_thick
 {
    border-right: solid 0px black !important;
	border-left: solid 3px black !important;
	border-top: solid 0px black !important;
	border-bottom: solid 1px black !important;
	
}


.no-top-right-border_Grey_thick
{
    border-right: solid 0px !important;
	border-left: solid 3px black !important;
	border-top: solid 0px !important;
	border-bottom: solid 1px #F4F4F4 !important; /* #EBEBEB */
/*
	border-top-color: red;
	border-bottom-color: red;
*/
}
.no-top-right-border_Grey
{
    border-right: solid 0px !important;
	border-left: solid 1px black !important;
	border-top: solid 0px !important;
	border-bottom: solid 1px #F4F4F4 !important;
/*
	border-top-color: red;
	border-bottom-color: red;
*/
}
.no-right-left-top-border_Grey
{
    border-right: solid 0px black !important;
	border-left: solid 0px black !important;
	border-top: solid 0px black !important;
	border-bottom: solid 1px #F4F4F4 !important;
	
}
.no-top-border_Grey
{
    border-right: solid 1px black !important;
	border-left: solid 1px black !important;
	border-top: solid 0px black !important;
	border-bottom: solid 1px #F4F4F4 !important;
	
	
}

.no-right-border_Grey_thick
{
    border-right: solid 0px black !important;
	border-left: solid 3px black !important;
	border-bottom: solid 1px black !important;
	border-top: solid 1px black !important;
	
}
.no-right-border_thick
{
    border-right: solid 0px black !important;
	border-left: solid 3px black !important;
	border-bottom: solid 1px black !important;
	border-top: solid 1px black !important;
	
}
.no-right-border {
    border-right: solid 0px black !important;
	border-left: solid 1px black !important;
	border-bottom: solid 1px black !important;
	border-top: solid 1px black !important;
	
}

.no-left-border {
    border-right: solid 1px black !important;
	border-left: solid 0px black !important;
	border-top: solid 1px black !important;
	border-bottom: solid 1px black !important;
	
}

.no-top-border {
    border-right: solid 1px black !important;
	border-left: solid 1px black !important;
	border-top: solid 0px black !important;
	border-bottom: solid 1px black !important;
	
	
}

.no-top-right-border {
    border-right: solid 0px black !important;
	border-left: solid 1px black !important;
	border-top: solid 0px black !important;
	border-bottom: solid 1px black !important;
	
}

.no-top-left-border {
    border-right: solid 1px black !important;
	border-left: solid 0px black !important;
	border-top: solid 0px black !important;
	border-bottom: solid 1px black !important;
	
}
.no-right-left-top-border {
    border-right: solid 0px black !important;
	border-left: solid 0px black !important;
	border-top: solid 0px black !important;
	border-bottom: solid 1px black !important;
	
}
.no-bottom-border {
    border-right: solid 1px black !important;
	border-left: solid 1px black !important;
	border-top: solid 1px black !important;
	border-bottom: solid 0px black !important;
}
.border-ALL_left_thick
{
    border-right: solid 1px black !important;
	border-left: solid 3px black !important;
	border-top: solid 1px black !important;
	border-bottom: solid 1px black !important;
	border-top-color:#E5E5E5;
}
.border-ALL {
    border-right: solid 1px black !important;
	border-left: solid 1px black !important;
	border-top: solid 1px black !important;
	border-bottom: solid 1px black !important;
	border-top-color:#E5E5E5;
}

.border-only-top-left-right-bottom {
    border-right: solid 0px black !important;
	border-left: solid 1px black !important;
	border-top: solid 1px black !important;
	border-bottom: solid 1px black !important;
	
}

.border-only-top-left-bottom {
    border-right: solid 1px black !important;
	border-left: solid 0px black !important;
	border-top: solid 1px black !important;
	border-bottom: solid 1px black !important;
}

	.AnnexFont
	{
		background-color:#FFFFFF;
		font-size:38pt;
		font-family:Arial;
		color:#000000;
		font-weight:bold;
	}

	.IndexHead1Font
	{
		background-color:#FFFFFF;
		font-size:14pt;
		font-family:Arial;
		color:#000000;
		font-weight:bold;
	}

	.IndexHead2Font
	{
		background-color:#D8D8D8;
		font-size:11pt;
		font-family:Arial;
		color:#000000;
		font-weight:bold;
	}
	
	.IndexContentFont
	{
		background-color:#FFFFFF;
		font-size:11pt;
		font-family:Arial;
		color:#000000;
		font-weight:bold;
	}


body {
  margin: 0;
  background: #CCCCCC;
}

div.page {

  margin-top: 1px ;
  margin-bottom: 5px ;
  margin-right: 10px ;
  margin-left: 10px ;

  /*margin: 10px auto;*/ /*10px*/


/*  border: solid 1px black;*/
  display: block;
  page-break-after: always;


   width: 209mm;
  height: 296mm;
 
 /* testing
   width: 209mm;
  height: 296mm;
 */
  overflow: hidden;
  background: white;
}

div.landscape-parent {
  width: 297mm;  /*296mm feb-2020*/
  height: 225mm; /*209mm feb-2020*/
}

div.landscape {
/*
	width: 296mm;
	height: 209mm;
*/
	width: 305mm; /*296mm feb-2020*/
	height: 217mm; /*209mm feb-2020*/
}

div.content {
 /* padding: 12mm; */
 /*padding-top: 2mm; *//* 13mm */
padding-bottom: 0mm;
 padding-right: 10mm;
 padding-left: 6mm;


}

body,
div,
td {
/*
  font-size: 10px;
  font-family: Arial;
*/ 
}

@media print {
  body {
    background: none;
  }
  div.page {
    width: 210mm; /* org 210mm */
    height: 300mm; /* org 296 */ 
  }
  div.landscape {
    transform: rotate(270deg) translate(-300mm, 0);
    transform-origin: 0 0;
  }
  div.portrait,
  div.landscape,
  div.page {
    margin: 0;
    padding: 0;
    border: none;
    background: none;
  }
}


div.applicant-break {page-break-after:always;}
/* This method works in IE6 */


	.AllReportHeader
	{
/*		border-width:1;*/
		border-style:none;
		background-color:#FFFFCA;
		font-size:11pt;

		font-family:Arial;
		font-weight:bold;

		color:#000000;
	}

	.AllReportHeaderSMALL
	{
/*		border-width:1;*/
		border-style:none;
		background-color:#FFFFCA;
		font-size:8pt;

		font-family:Arial;
		font-weight:bold;

		color:#000000;
	}

	.pageNumberAtPostion
	{
		font-family: Arial;
		font-size: 8pt;
		color:#000000;
		font-weight: bold;
		
		position: relative;
		top:0px;
		width: 10%;
		align-items:center;
		
		/*
		position: fixed;
		bottom: 0;
		*/
	}

	.pageNumber
	{
		font-family: Arial;
		font-size: 9pt;
		color:#000000;
		font-weight: normal;
	}

	.RepfontStylePage17SpecialColor
	{
		background-color:#395682;

		font-family:Arial;
		font-size:12pt;

		font-weight:bold;
		color:#FFFFFF;
	}

	.RepfontStylePage17
	{
		background-color:#FFFFCA;

		font-family:Arial;
		font-size:12pt;

		font-weight:bold;
		color:#000000;
	}

	.RepfontStylePage14
	{
		background-color:#FFFFCA;

		font-family:Arial;
		font-size:12pt;

		font-weight:bold;
		color:#000000;
	}



	.RepfontStylePage8
	{
		background-color:#FFFFCA;

		font-family:Arial;
		font-size:11pt;

		font-weight:bold;
		color:#000000;
	}

	.RepfontStyle
	{
		background-color:#FFFFCA;

		font-family:Arial;
		font-size:10pt;

		font-weight:bold;
		color:#000000;
	}
	.RepfontStylePage7
	{
		background-color:#FFFFCA;

		font-family:Arial;
		font-size:11pt;

		font-weight:bold;
		color:#000000;
	}
	.RepfontStylePrintReport
	{
		background-color:#FFFFCA;

		font-family:Arial;
		font-size:10pt;

		font-weight:bold;
		color:#000000;
	}

	.RepfontStyleSMALL
	{
		border-style:none;
		background-color:#FFFFCA;
		font-family:Arial;
		font-size:8pt;
		font-weight:bold;
		color:#000000;
	}

	.RepfontStylePage15And16
	{
		border-style:none;
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:9pt;
		font-weight:normal;
		color:#000000;
	}


	.RepfontStyleCalculatedFiledPage15And16
	{
		background-color:#E4E4E4; /*#D6D6D6;*/
		font-family:Arial;
		font-size:9pt;
		font-weight:bold;
		color:#000000;
	}

	.SubRepHeaderPage15And16
	{
		
		background-color:#FFFFCA;
		font-family:Arial;
		font-size:8pt;
		font-weight:normal;
		color:#000000;
	}

	.SubRepHeaderPage15And16Calculated
	{
		border-style:normal;
		background-color:#B8C7E2;
		font-family:Arial;
		font-size:8pt;
	}

	.RepHeader55555
	{
/*		border-width:0;*/
		border-style:none;
		background-color:#FFFFCA;
		font-family:Arial;
		font-size:10pt;
		font-weight:bold;
		color:#000000;
	}


	.RepContentFontRed
	{
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:9pt;
		color:#BB0428;
		font-weight:bold;
	}

	
	.RepContentFont
	{
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:10pt;
		color:#000000;
		font-weight:bold; 
	}

	.RepContentFontNormal
	{
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:10pt;
		color:#000000;
		font-weight:normal; 
	}

	

	.RepContentFontNormalPage8
	{
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:11pt;
		color:#000000;
		font-weight:normal; 
	}
	.RepContentFontNormalPage7
	{
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:11pt;
		color:#000000;
		font-weight:normal; 
	}

	.RepContentFontPage17
	{
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:11pt;
		color:#000000;
		font-weight:normal;
	}

	.RepfontStyleCalculatedFiledPage17
	{
		background-color:#E4E4E4; /*#D6D6D6;*/
		font-family:Arial;
		font-size:11pt;
		font-weight:normal;
		color:#000000;
	}

	.RepContentFontPrintReport
	{
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:10pt;
		color:#000000;
		font-weight:bold; /*bold; feb-2020*/
	}

	.RepContentFontPage14
	{
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:12pt;
		color:#000000;
		font-weight:bold;
	}

	.RepContentFontNormalSMALL
	{
	/*	padding-left:100px;
		padding-top:100px;
		padding-right:100px;
		padding-bottom:100px;*/
		/*padding: 120px 135px 10px 10px;
		background-color:#FFFFFF;
	*/
		font-family:Arial;
		font-size:9pt;
		color:black;
		font-weight:normal; /*bold; feb-2020*/
	}


	.RepContentFontSMALL
	{
	/*	padding-left:100px;
		padding-top:100px;
		padding-right:100px;
		padding-bottom:100px;*/
		/*padding: 120px 135px 10px 10px;
		background-color:#FFFFFF;
	*/
		font-family:Arial;
		font-size:9pt;
		color:black;
		font-weight:bold; 
	}

	.RepfontStyleReportRemark
	{
		font-style:italic;
		background-color:#FFFFFF;
		font-family:Arial;
		font-size:10pt;
		color:#B30000;
		font-weight:normal;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}


	.RepfontStyleReportRemarkSMALL
	{
		font-style:italic;
		background-color:#FFFFFF;

		font-family:Arial;
		font-size:8pt;

		color:#FF0000;
		font-weight:normal;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}


	.RepfontStyleItalicPrintReport
	{
		font-style:italic;
		background-color:#FFFFFF;

		font-family:Arial;
		font-size:8pt;

		font-weight:normal;
		color:purple;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}

	.RepfontStyleItalic
	{
		font-style:italic;
		background-color:#FFFFFF;

		font-family:Arial;
		font-size:9pt;

		font-weight:bold;
		color:#0033FF;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}





	.RepfontStyleCalculatedFiledPrintReportPage7
	{
	
		background-color:#E4E4E4; /*#D6D6D6;*/

		font-family:Arial;
		font-size:12pt;
		font-weight:bold;
		color:#000000;
	}

	.RepfontStyleCalculatedFiled
	{
	
		background-color:#E4E4E4; /*#D6D6D6;*/
		font-family:Arial;
		font-size:10pt;
		font-weight:bold;
		color:#000000;
	}

	.RepfontStyleCalculatedFiledPage8
	{
	
		background-color:#E4E4E4; /*#D6D6D6;*/
		font-family:Arial;
		font-size:11pt;
		font-weight:bold;
		color:#000000;
	}

	.RepfontStyleCalculatedFiledSMALL
	{
	
		background-color:#E4E4E4; /*#D6D6D6;*/

		font-family:Arial;
		font-size:9pt;
		font-weight:bold;
		color:#000000;
	}

	.RepfontStyleCalculatedFiledPage14
	{
	
		background-color:#E4E4E4; /*#D6D6D6;*/

		font-family:Arial;
		font-size:10pt;
		font-weight:bold;
		color:#000000;
	}


	.RepfontStyleTransparentHeading
	{
/*		border-width:1;*/
		border-style:none;
/*
		font-family:Arial;
		font-size:13pt;
*/
		font-weight:bold;
		color:#000000;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}


	.RepfontStyleMainTransparentHeading
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:13pt;

		font-weight:bold;
		color:#000000;
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
	}

	.RepfontStyleCommonTransparentHeadingSmall
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:11pt;

		font-weight:bold;
		color:#000000;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}
	.RepfontStyleCommonTransparentHeading
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:11pt;

		font-weight:bold;
		color:#000000;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}


	.GreenTransparentHeadingVerySmall
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:8pt;

		font-weight:bold;
		color:#075256;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}

	.GreenTransparentHeadingSmall
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:11pt;

		font-weight:bold;
		color:#075256;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}

.GreenTransparentHeadingNormal
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:12pt;

		font-weight:bold;
		color:#075256;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}

.GreenTransparentHeadingNormalPage7
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:12pt;

		font-weight:bold;
		color:#075256;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}

	
	.GreenTop10Page7
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:12pt;

		font-weight:bold;
		color:Green;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}

	.RedTransparentHeadingSmallPage7
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:12pt;

		font-weight:bold;
		color:#B90000;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}

	.GreenTransparentHeading
	{
/*		border-width:1;*/
		border-style:none;

		font-family:Arial;
		font-size:13pt;

		font-weight:bold;
		color:#075256;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}


	.RepfontStyle1
	{
/*		border-width:0;*/
		border-style:none;
		background-color:#FFFFFF;

/*
		font-family:Arial;
		font-size:11pt;
*/
		font-weight:bold;
		color:#000000;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}


	.cellStyle1
	{
		background-color:#f9fddf;
/*
		font-family: Arial;
		font-size: 11pt;
*/
		color:#000000;
		font-weight: bold;
/*
		padding-left:10;
		padding-top:10;
		padding-right:10;
		padding-bottom:10;
*/
	}


</STYLE>

	<Body Leftmargin=5>
	<Form Method="Post" Name="form">
	<center>
	<Script Src="Common.js"></Script>	
	<Script Language="javascript">
	function printReport(form)
	{
		printDiv.style.visibility="hidden";
		window.document.execCommand("Print");
		printDiv.style.visibility="visible";
		return false;
	}
	</Script>
	<!--
	<Center>
	-->
	<%
		If Trim(Request.Form("optPrint")) = "V" Then
	%>
	<!--
		<Div id=printDiv class=divStyle>
		<A Href=# onclick="return printReport(form)"><Img Src='\images\Print.jpg' Border=0></A>
		</Div>
	-->
		<P><BR></P>
	<%
		End if
	%>
	<!--
	style="position:absolute;Top:1;left:10;height:40;" 
	<Table class="spanStyle" Cellpadding=0 width=650 Cellspacing=1 border=0>
	-->

	<%

response.write "<div class='page'>"
response.write "<div class='content'>"
response.write "<Table border=0 cellPadding=1 cellspacing=0 height='95%'  width='100%' >"
response.write "<TR >"
response.write "<TD valign=top>"


	call CoverPage
response.write "</TD>"
response.write "</TR>"

response.write "</Table>"

response.write "</div>"
response.write "</div>"


response.write "<div class='page'>"
response.write "<div class='content'>"
response.write "<Table border=0 cellPadding=1 cellspacing=0 height='95%'  width='100%' >"
response.write "<TR >"
response.write "<TD valign=top>"


	call IndexPrint
response.write "</TD>"
response.write "</TR>"

response.write "</Table>"

response.write "</div>"
response.write "</div>"



	response.write "<div class='page'>"
	response.write "<div class='content'>"
	response.write "<Table border=0 cellPadding=0 cellspacing=0 height='98%' width='100%' >"
	response.write "<TR >"
	response.write "<TD valign=top>"

	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleMainTransparentHeading><Td >" 
	sSql = sSql & "Group Results for " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "&nbsp;"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 

	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 


	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Financial Highlights"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 

	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Group Level"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

	sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_MTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_MTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL) VAR1,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL)/SUM(DECODE(BGT_MTD_BAL,0,1,BGT_MTD_BAL)) *100 VARP,"
	sSql = sSql & " SUM(ACT_MTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1) VAR "
	'sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1)/SUM(DECODE(ACT_MTD_BAL_PY1,0,1,ACT_MTD_BAL_PY1))*100 VARLY"
	'sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1)/SUM(decode(nvl(ACT_MTD_BAL_PY1,1),1,1 ,nvl(ACT_MTD_BAL_PY1,1)))*100 VARLY "

	'sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1)/SUM(decode(ACT_MTD_BAL_PY1,0,1 ,nvl(ACT_MTD_BAL_PY1,0),1,ACT_MTD_BAL_PY1))*100 VARLY " ' given by cherian

	sSql = sSql & " from mistxn"
	sSql = sSql & " where"
	sSql = sSql & " cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' AND"
	sSql = sSql & " CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"

'response.write sSql
'response.end
	tblHeading="Month "  & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	'

	response.write "<br>"



	call printReport(sSql,tblHeading)
	'
	response.write "<br>"
	'
	sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_YTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_YTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL) VAR1,"
	sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL)/SUM(DECODE(BGT_YTD_BAL,0,1,BGT_YTD_BAL))*100 VARP,"
	sSql = sSql & " SUM(ACT_YTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_YTD_BAL-ACT_YTD_BAL_PY1) VAR,"
	sSql = sSql & " SUM(ACT_YTD_BAL-ACT_YTD_BAL_PY1)/SUM(DECODE(ACT_YTD_BAL_PY1,0,1,ACT_YTD_BAL_PY1)) *100 VARLY"
	sSql = sSql & " from mistxn"
	sSql = sSql & " where"
	sSql = sSql & " cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' AND"
	sSql = sSql & " CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"
	'

'response.write sSql
'response.end
	tblHeading="YTD "  & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	'
	call printReport(sSql,tblHeading)
	'
	
	response.write "<br>"

	'
	call Summary5Years()
	'
response.write "</TD>"
response.write "</TR>"
response.write "</Table>"
response.write "<font class='pageNumber'>Page 1<font>"  

'response.write "<div class='pageNumberAtPostion'>"
'response.write "<font>Page 1<font>"    
'response.write "</div >"


response.write "</div>"
response.write "</div>"
	
response.write "<div class='page'>"
response.write "<div class='content'>"
response.write "<Table border=0 cellPadding=1 cellspacing=0 height='98%' width='100%' >"
response.write "<TR >"
response.write "<TD valign=top>"

sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 

	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Revenue by Company"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

	call page2()
response.write "</TD>"
response.write "</TR>"


response.write "</Table>"
response.write "<font class='pageNumber'>Page 2<font>"  
response.write "</div>"
response.write "</div>"
	
response.write "<div class='page'>"
response.write "<div class='content'>"
response.write "<Table border=0 height='98%' width='100%' >"
response.write "<TR>"
response.write "<TD valign=top>"

sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 

	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "PBT-Operations by Company (excl. Investment Income)"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

	sSql=" select "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL,0)) FLASH , "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL_PY1,0)) LY, "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL,0)) - SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL_PY1,0)) VS_LY, "
	sSql= sSql & " B.ccompnamesn2, "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL,0)) YearlyFLASH , "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL_PY1,0)) YearlyLY, "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL,0)) - SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL_PY1,0)) YearlyVS_LY "
	sSql= sSql & " from mistxn A , COMPANY B "
	sSql= sSql & " WHERE A.CCOMPCODE = B.CCOMPCODE "
	sSql= sSql & " AND cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
	'sSql= sSql & " AND A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	'sSql= sSql & " AND A.CCOMPCODE IN ('AVOD','AMIANTIT','AMO','IMOS','SPOL','AMLSC','RBP','NTS','JSS','OTM','ZTC') "
	sSql= sSql & " AND B.sectioninBBReport ='A' "
	sSql= sSql & " AND DESCCD IN ('R5') "
	sSql= sSql & " GROUP BY CYEAR, CMONTH,B.ccompnamesn2, CATEGORY "
	sSql= sSql & " ORDER BY 5 desc "

	SectionATotalCol1=0
	SectionATotalCol2=0
	SectionATotalCol3=0
	SectionATotalCol4=0
	SectionATotalCol5=0
	SectionATotalCol6=0

	call page3("Y")

	sSql=" select "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL,0)) FLASH , "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL_PY1,0)) LY, "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL,0)) - SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL_PY1,0)) VS_LY, "
	sSql= sSql & " B.ccompnamesn2, "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL,0)) YearlyFLASH , "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL_PY1,0)) YearlyLY, "
	sSql= sSql & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL,0)) - SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL_PY1,0)) YearlyVS_LY "
	sSql= sSql & " from mistxn A , COMPANY B "
	sSql= sSql & " WHERE A.CCOMPCODE = B.CCOMPCODE "
	sSql= sSql & " AND cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
	'sSql= sSql & " AND A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC','AVOD','AMIANTIT','AMO','IMOS','SPOL','AMLSC','RBP','NTS','JSS','OTM','ZTC') "
	sSql= sSql & " AND B.sectioninBBReport ='B' "
	sSql= sSql & " AND DESCCD IN ('R5') "
	sSql= sSql & " GROUP BY CYEAR, CMONTH,B.ccompnamesn2, CATEGORY "
	sSql= sSql & " ORDER BY 5 desc "

	call page3("N")
response.write "</TD>"
response.write "</TR>"


response.write "</Table>"
response.write "<font class='pageNumber'>Page 3<font>"    

response.write "</div>"
response.write "</div>"
	
response.write "<div class='page'>"
response.write "<div class='content'>"
response.write "<Table border=0 height='98%' width='100%' >"
response.write "<TR >"
response.write "<TD valign=top>"

sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 

	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Cash PBT-Operations by Company (excl. Investment Income)"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

	call page4()

response.write "</TD>"
response.write "</TR>"


response.write "</Table>"
response.write "<font class='pageNumber'>Page 4<font>"    
response.write "</div>"
response.write "</div>"
	



'response.write "<div class='page landscape-parent'>"
'response.write "<div class='landscape'>"
'response.write "<div class='content'>"
'response.write "<Table border=0 width='95%' height='100%'>"
'response.write "<TR >"
'response.write "<TD valign=top>"

	

response.write "<div class='page'>"
response.write "<div class='content'>"
response.write "<Table border=0 height='98%' width='100%' >"
response.write "<TR >"
response.write "<TD valign=top>"
	'
	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 


	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Financial Highlights"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 

	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Manufacturing Sector"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

	sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_MTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_MTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL) VAR1,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL)/SUM(DECODE(BGT_MTD_BAL,0,1,BGT_MTD_BAL))*100 VARP,"
	sSql = sSql & " SUM(ACT_MTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1) VAR,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1)/SUM(DECODE(ACT_MTD_BAL_PY1,0,1,ACT_MTD_BAL_PY1))*100 VARLY"
	sSql = sSql & " from mistxn  A , COMPANY B"
	sSql = sSql & " where"
	sSql = sSql & " A.CCOMPCODE = B.CCOMPCODE "
	sSql = sSql & " And cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
	sSql = sSql & " And A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND B.CATEGORY ='Manufacturing'"
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"
	'
	'response.write ssql

'response.write sSql
'response.end

	tblHeading="Month "  & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	'
	call printReport(sSql,tblHeading)
	'
	response.write "<br>"
	'
	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Trading Sector and Contracting Sector"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

	sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_MTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_MTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL) VAR1,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL)/SUM(DECODE(BGT_MTD_BAL,0,1,BGT_MTD_BAL))*100 VARP,"
	sSql = sSql & " SUM(ACT_MTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1) VAR,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1)/SUM(DECODE(ACT_MTD_BAL_PY1,0,1,ACT_MTD_BAL_PY1))*100 VARLY"
	sSql = sSql & " from mistxn  A , COMPANY B"
	sSql = sSql & " where"
	sSql = sSql & " A.CCOMPCODE = B.CCOMPCODE "
	sSql = sSql & " And cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
	sSql = sSql & " And A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND B.CATEGORY in ('Trading','Contracting') "
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"
	'
	tblHeading="Month  "  & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	'
'response.write sSql	
'response.end
	call printReport(sSql,tblHeading)
	'
	response.write "<br>"
	'
	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Service and Investment Sector"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

	sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_MTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_MTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL) VAR1,"
	'sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL)/SUM(DECODE(BGT_MTD_BAL,0,1,BGT_MTD_BAL))*100 VARP,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL)/decode(SUM(DECODE(nvl(BGT_MTD_BAL,0),0,1,BGT_MTD_BAL)),0,1,SUM(DECODE(nvl(BGT_MTD_BAL,0),0,1,BGT_MTD_BAL)))*100 VARP, "
	sSql = sSql & " SUM(ACT_MTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1) VAR,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1)/SUM(DECODE(ACT_MTD_BAL_PY1,0,1,ACT_MTD_BAL_PY1))*100 VARLY"
	sSql = sSql & " from mistxn  A , COMPANY B"
	sSql = sSql & " where"
	sSql = sSql & " A.CCOMPCODE = B.CCOMPCODE "
	sSql = sSql & " And cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
	sSql = sSql & " And A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND B.CATEGORY like 'Service%' "
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"
	'
'response.write sSql	
'response.end
	tblHeading="Month "  & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	'
	call printReport(sSql,tblHeading)
	'
response.write "</TD>"
response.write "</TR>"


response.write "</Table>"
response.write "<font class='pageNumber'>Page 5<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"







'response.write "<div class='page landscape-parent'>"
'response.write "<div class='landscape'>"
'response.write "<div class='content'>"
'response.write "<Table border=0 width='100%' height='100%'>"
'response.write "<TR height='95%' >"
'response.write "<TD valign=top>"

response.write "<div class='page'>"
response.write "<div class='content'>"
response.write "<Table border=0 height='98%' width='100%' >"
response.write "<TR >"
response.write "<TD valign=top>"
	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 


	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Financial Highlights"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 

	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Manufacturing Sector"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	
	sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_YTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_YTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL) VAR1,"
	sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL)/SUM(DECODE(BGT_YTD_BAL,0,1,BGT_YTD_BAL))*100 VARP,"
	sSql = sSql & " SUM(ACT_YTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_YTD_BAL-ACT_YTD_BAL_PY1) VAR,"
	sSql = sSql & " SUM(ACT_YTD_BAL-ACT_YTD_BAL_PY1)/SUM(DECODE(ACT_YTD_BAL_PY1,0,1,ACT_YTD_BAL_PY1))*100 VARLY"
	sSql = sSql & " from mistxn  A , COMPANY B"
	sSql = sSql & " where"
	sSql = sSql & " A.CCOMPCODE = B.CCOMPCODE "
	sSql = sSql & " And cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
	sSql = sSql & " And A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND B.CATEGORY ='Manufacturing'"
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"
	'
	tblHeading="YTD "  & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	'
'	response.write sSql
'	response.end
	call printReport(sSql,tblHeading)
	'
	response.write "<br>"
	'
	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Trading Sector and Contracting Sector"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	
	sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_YTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_YTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL) VAR1,"
	sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL)/SUM(DECODE(BGT_YTD_BAL,0,1,BGT_YTD_BAL))*100 VARP,"
	sSql = sSql & " SUM(ACT_YTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_YTD_BAL-ACT_YTD_BAL_PY1) VAR,"
	sSql = sSql & " SUM(ACT_YTD_BAL-ACT_YTD_BAL_PY1)/SUM(DECODE(ACT_YTD_BAL_PY1,0,1,ACT_YTD_BAL_PY1))*100 VARLY"
	sSql = sSql & " from mistxn  A , COMPANY B"
	sSql = sSql & " where"
	sSql = sSql & " A.CCOMPCODE = B.CCOMPCODE "
	sSql = sSql & " And cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
	sSql = sSql & " And A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND B.CATEGORY in ('Trading','Contracting')"
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"
	'
	tblHeading="YTD "  & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	'
'	response.write sSql
'	response.end
	call printReport(sSql,tblHeading)
	'
	response.write "<br>"
	'
	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=GreenTransparentHeading><Td>" 
	sSql = sSql & "Service and Investment Sector"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

	sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_YTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_YTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL) VAR1,"
	'sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL)/SUM(DECODE(BGT_YTD_BAL,0,1,BGT_YTD_BAL))*100 VARP,"

	sSql = sSql & " SUM(ACT_YTD_BAL-BGT_YTD_BAL)/decode(SUM(DECODE(nvl(BGT_YTD_BAL,0),0,1,BGT_YTD_BAL)),0,1,SUM(DECODE(nvl(BGT_YTD_BAL,0),0,1,BGT_YTD_BAL)))*100 VARP, "

	sSql = sSql & " SUM(ACT_YTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_YTD_BAL-ACT_YTD_BAL_PY1) VAR,"
	sSql = sSql & " SUM(ACT_YTD_BAL-ACT_YTD_BAL_PY1)/SUM(DECODE(ACT_YTD_BAL_PY1,0,1,ACT_YTD_BAL_PY1))*100 VARLY"
	sSql = sSql & " from mistxn  A , COMPANY B"
	sSql = sSql & " where"
	sSql = sSql & " A.CCOMPCODE = B.CCOMPCODE "
	sSql = sSql & " And cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
	sSql = sSql & " And A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND B.CATEGORY like 'Service%'"
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"
	'
	tblHeading="YTD "  & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear
	'
'	response.write ssql
'	response.end
	call printReport(sSql,tblHeading)

response.write "</TD>"
response.write "</TR>"


response.write "</Table>"
response.write "<font class='pageNumber'>Page 6<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"



	response.write "<div class='page'>"
	response.write "<div class='content'>"
response.write "<Table border=0 height='98%' width='100%' >"
response.write "<TR >"
response.write "<TD valign=top>"

	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 


	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Major contributions to Profit/(Loss) from Operations (excl. Investment Income)"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "&nbsp;"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 

	sSql = sSql & "</Table>"
	response.write sSql	
	
	response.write "<Table border=0 height='91%' width='100%' >"
	response.write "<TR >"
	response.write "<TD valign=top>"




call page7_Top10_Bottom10()



response.write "</TD>"
response.write "</TR>"


response.write "</Table>"


response.write "</TD>"
response.write "</TR>"


response.write "</Table>"

response.write "<font class='pageNumber'>Page 7<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"



response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"
response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR >"
response.write "<TD valign=top>"
	


	sSql1=  " select c.ccompnamesn2,a.desccd,"
	sSql1= sSql1 & " sum(a.act_mtd_bal) Flash, "
	sSql1= sSql1 & " sum(a.bgt_mtd_bal) Budget "
	sSql1= sSql1 & " from mistxn a, Company c "
	sSql1= sSql1 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql1= sSql1 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql1= sSql1 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql1= sSql1 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql1= sSql1 & " AND CATEGORY ='Manufacturing' "
	sSql1= sSql1 & " group by c.ccompnamesn2,a.desccd "
	sSql1= sSql1 & " order by 1,2 "

	sSql2=  " select c.ccompnamesn2,a.desccd,"
	sSql2= sSql2 & " sum(a.act_mtd_bal) Flash, "
	sSql2= sSql2 & " sum(a.act_mtd_bal_PY1) Budget "
	sSql2= sSql2 & " from mistxn a, Company c "
	sSql2= sSql2 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql2= sSql2 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql2= sSql2 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql2= sSql2 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql2= sSql2 & " AND CATEGORY ='Manufacturing' "
	sSql2= sSql2 & " group by c.ccompnamesn2,a.desccd "
	sSql2= sSql2 & " order by 1,2"

	Title1="MTD Variance Analysis versus Budget - Manufacturing Sector"
	Title2="MTD Variance Analysis versus Last Year - Manufacturing Sector"

	call page8_9_10_11_12_13(sSql1,sSql2,Title1,Title2,"M","NORMAL")
	'
response.write "</TD>"
response.write "</TR>"

'response.write "<TR>"
'response.write "<TD align=center>"

'response.write "</TD>"
'response.write "</TR>"
   
response.write "</Table>"
response.write "<font class='pageNumber'>Page 8<font>" 
response.write "</div>"
response.write "</div>"
response.write "</div>"









response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"
response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR >"
response.write "<TD valign=top>"


	'sSql1=  " select a.ccompcode,a.desccd,"
	sSql1=  " select c.ccompnamesn2,a.desccd,"
	sSql1= sSql1 & " sum(a.act_mtd_bal) Flash, "
	sSql1= sSql1 & " sum(a.bgt_mtd_bal) Budget "
	sSql1= sSql1 & " from mistxn a, Company c "
	sSql1= sSql1 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql1= sSql1 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql1= sSql1 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql1= sSql1 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql1= sSql1 & " AND CATEGORY in ('Trading','Contracting') "
	sSql1= sSql1 & " group by c.ccompnamesn2,a.desccd "
	sSql1= sSql1 & " order by 1,2 "

	'sSql2=  " select a.ccompcode,a.desccd,"
	sSql2=  " select c.ccompnamesn2,a.desccd,"
	sSql2= sSql2 & " sum(a.act_mtd_bal) Flash, "
	sSql2= sSql2 & " sum(a.act_mtd_bal_PY1) Budget "
	sSql2= sSql2 & " from mistxn a, Company c "
	sSql2= sSql2 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql2= sSql2 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql2= sSql2 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql2= sSql2 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql2= sSql2 & " AND CATEGORY in ('Trading','Contracting') "
	sSql2= sSql2 & " group by c.ccompnamesn2,a.desccd "
	sSql2= sSql2 & " order by 1,2 "

	Title1="MTD Variance Analysis versus Budget -  Trading Sector & Contracting Sector"
	Title2="MTD Variance Analysis versus Last Year -  Trading Sector & Contracting Sector"



	call page8_9_10_11_12_13(sSql1,sSql2,Title1,Title2,"M","NORMAL")
response.write "</TD>"
response.write "</TR>"

'response.write "<TR>"
'response.write "<TD align=center>"

'response.write "</TD>"
'response.write "</TR>"

response.write "</Table>"
response.write "<font class='pageNumber'>Page 9<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"









response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"
response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR  >"
response.write "<TD valign=top>"

	'sSql1=  " select a.ccompcode,a.desccd,"
	sSql1=  " select c.ccompnamesn2,a.desccd,"
	sSql1= sSql1 & " sum(a.act_mtd_bal) Flash, "
	sSql1= sSql1 & " sum(a.bgt_mtd_bal) Budget "
	sSql1= sSql1 & " from mistxn a, Company c "
	sSql1= sSql1 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql1= sSql1 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql1= sSql1 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql1= sSql1 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql1= sSql1 & " AND CATEGORY like 'Service%' "
	sSql1= sSql1 & " group by c.ccompnamesn2,a.desccd "
	sSql1= sSql1 & " order by 1,2"

	'sSql2=  " select a.ccompcode,a.desccd,"
	sSql2=  " select c.ccompnamesn2,a.desccd,"
	sSql2= sSql2 & " sum(a.act_mtd_bal) Flash, "
	sSql2= sSql2 & " sum(a.act_mtd_bal_PY1) Budget "
	sSql2= sSql2 & " from mistxn a, Company c "
	sSql2= sSql2 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql2= sSql2 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql2= sSql2 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql2= sSql2 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql2= sSql2 & " AND CATEGORY like 'Service%' "
	sSql2= sSql2 & " group by c.ccompnamesn2,a.desccd "
	sSql2= sSql2 & " order by 1,2 "

	Title1="MTD Variance Analysis versus Budget -  Service and Investment Sector"
	Title2="MTD Variance Analysis versus Last Year -  Service and Investment Sector"


		call page8_9_10_11_12_13(sSql1,sSql2,Title1,Title2,"M","SMALL")
response.write "</TD>"
response.write "</TR>"

'response.write "<TR>"
'response.write "<TD align=center>"

'response.write "</TD>"
'response.write "</TR>"

response.write "</Table>"
response.write "<font class='pageNumber'>Page 10<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"






response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"
response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR>"
response.write "<TD valign=top>"

	'sSql1=  " select a.ccompcode,a.desccd,"
	sSql1=  " select c.ccompnamesn2,a.desccd,"
	sSql1= sSql1 & " sum(a.act_ytd_bal) Flash, "
	sSql1= sSql1 & " sum(a.bgt_ytd_bal) Budget "
	sSql1= sSql1 & " from mistxn a, Company c "
	sSql1= sSql1 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql1= sSql1 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql1= sSql1 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql1= sSql1 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql1= sSql1 & " AND CATEGORY ='Manufacturing' "
	sSql1= sSql1 & " group by c.ccompnamesn2,a.desccd "
	sSql1= sSql1 & " order by 1,2 "

	'sSql2=  " select a.ccompcode,a.desccd,"
	sSql2=  " select c.ccompnamesn2,a.desccd,"
	sSql2= sSql2 & " sum(a.act_ytd_bal) Flash, "
	sSql2= sSql2 & " sum(a.act_Ytd_bal_PY1) Budget "
	sSql2= sSql2 & " from mistxn a, Company c "
	sSql2= sSql2 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql2= sSql2 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql2= sSql2 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql2= sSql2 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql2= sSql2 & " AND CATEGORY ='Manufacturing' "
	sSql2= sSql2 & " group by c.ccompnamesn2,a.desccd "
	sSql2= sSql2 & " order by 1,2 "

	Title1="YTD Variance Analysis versus Budget - Manufacturing Sector"
	Title2="YTD Variance Analysis versus Last Year - Manufacturing Sector"

		call page8_9_10_11_12_13(sSql1,sSql2,Title1,Title2,"Y","NORMAL")
response.write "</TD>"
response.write "</TR>"

'response.write "<TR>"
'response.write "<TD align=center>"

'response.write "</TD>"
'response.write "</TR>"

response.write "</Table>"
response.write "<font class='pageNumber'>Page 11<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"





response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"
response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR>"
response.write "<TD valign=top>"

	'sSql1=  " select a.ccompcode,a.desccd,"
	sSql1=  " select c.ccompnamesn2,a.desccd,"
	sSql1= sSql1 & " sum(a.act_ytd_bal) Flash, "
	sSql1= sSql1 & " sum(a.bgt_ytd_bal) Budget "
	sSql1= sSql1 & " from mistxn a, Company c "
	sSql1= sSql1 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql1= sSql1 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql1= sSql1 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql1= sSql1 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql1= sSql1 & " AND CATEGORY in ('Trading','Contracting') "
	sSql1= sSql1 & " group by c.ccompnamesn2,a.desccd "
	sSql1= sSql1 & " order by 1,2 "

	'sSql2=  " select a.ccompcode,a.desccd,"
	sSql2=  " select c.ccompnamesn2,a.desccd,"
	sSql2= sSql2 & " sum(a.act_ytd_bal) Flash, "
	sSql2= sSql2 & " sum(a.act_Ytd_bal_PY1) Budget "
	sSql2= sSql2 & " from mistxn a, Company c "
	sSql2= sSql2 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql2= sSql2 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql2= sSql2 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql2= sSql2 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql2= sSql2 & " AND CATEGORY in ('Trading','Contracting') "
	sSql2= sSql2 & " group by c.ccompnamesn2,a.desccd "
	sSql2= sSql2 & " order by 1,2 "

	Title1="YTD Variance Analysis versus Budget - Trading Sector & Contracting Sector"
	Title2="YTD Variance Analysis versus Last Year - Trading Sector & Contracting Sector"

			call page8_9_10_11_12_13(sSql1,sSql2,Title1,Title2,"Y","NORMAL")
response.write "</TD>"
response.write "</TR>"

'response.write "<TR>"
'response.write "<TD align=center>"

'response.write "</TD>"
'response.write "</TR>"

response.write "</Table>"
response.write "<font class='pageNumber'>Page 12<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"



response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"
response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR >"
response.write "<TD valign=top>"

	'sSql1=  " select a.ccompcode,a.desccd,"
	sSql1=  " select c.ccompnamesn2,a.desccd,"
	sSql1= sSql1 & " sum(a.act_ytd_bal) Flash, "
	sSql1= sSql1 & " sum(a.bgt_ytd_bal) Budget "
	sSql1= sSql1 & " from mistxn a, Company c "
	sSql1= sSql1 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql1= sSql1 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql1= sSql1 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql1= sSql1 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql1= sSql1 & " AND  CATEGORY like 'Service%' "
	sSql1= sSql1 & " group by c.ccompnamesn2,a.desccd "
	sSql1= sSql1 & " order by 1,2 "

	'sSql2=  " select a.ccompcode,a.desccd,"
	sSql2=  " select c.ccompnamesn2,a.desccd,"
	sSql2= sSql2 & " sum(a.act_ytd_bal) Flash, "
	sSql2= sSql2 & " sum(a.act_Ytd_bal_PY1) Budget "
	sSql2= sSql2 & " from mistxn a, Company c "
	sSql2= sSql2 & " WHERE A.CCOMPCODE = C.CCOMPCODE "
	sSql2= sSql2 & " and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "'  "
	sSql2= sSql2 & " and a.desccd in ('Q5','R5','E5','K5') "
	sSql2= sSql2 & " AND C.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
	sSql2= sSql2 & " AND  CATEGORY like 'Service%' "
	sSql2= sSql2 & " group by c.ccompnamesn2,a.desccd "
	sSql2= sSql2 & " order by 1,2 "

	Title1="YTD Variance Analysis versus Budget - Service and Investment Sector"
	Title2="YTD Variance Analysis versus Last Year - Service and Investment Sector"
	'
	call page8_9_10_11_12_13(sSql1,sSql2,Title1,Title2,"Y","SMALL")	
response.write "</TD>"
response.write "</TR>"

'response.write "<TR>"
'response.write "<TD align=center>"
 
'response.write "</TD>"
'response.write "</TR>"

response.write "</Table>"
response.write "<font class='pageNumber'>Page 13<font>"   
response.write "</div>"
response.write "</div>"
response.write "</div>"
''

	response.write "<div class='page'>"
	response.write "<div class='content'>"

'response.write "<div class='page landscape-parent'>"
'response.write "<div class='landscape'>"
'response.write "<div class='content'>"


	sSql =  "<Table border=0  width='99%'   cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "PBT for ZTC & OFCC"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "<TR Class=GreenTransparentHeading><Td >" 
	sSql = sSql & "Zawawi Trading Company (ZTC)" 
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
	response.write sSql	

response.write "<Table border=0 width='99%' height='92%'>"
response.write "<TR >"
response.write "<TD valign=top>"


	call Page14()



response.write "</Table>"
response.write "<font class='pageNumber'>Page 14<font>"
response.write "</div>"
response.write "</div>"
response.write "</div>"


''
response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"

'	sSql =  "<Table border=0  width='95%' cellPadding=0 cellspacing=0	>"
'	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeadingSmall><Td >" 
'	sSql = sSql & "Amount in RO '000"
'	sSql = sSql & "</Td >" 
'	sSql = sSql & "</TR>" 
'	sSql = sSql & "<TR Class=GreenTransparentHeadingVerySmall><Td>" 
'	sSql = sSql & "Company-wise Performance Summary of Group Results for the period:&nbsp;" & FullMonthsArr(cint(TheReportMonth)-1) & "/" & TheReportYear & "&nbsp;(MTD)"
'	sSql = sSql & "</Td >" 
'	sSql = sSql & "</TR>" 
'	sSql = sSql & "</Table>"
'	response.write sSql	


response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR >"
response.write "<TD valign=top align=left>"


call Page15("M")

'response.write "<TR>"
'response.write "<TD align=center>"

'response.write "</TD>"
'response.write "</TR>"

response.write "</Table>"
response.write "<font class='pageNumber'>Page 15<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"


response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"

'	sSql =  "<Table border=0  width='95%' cellPadding=1 cellspacing=0>"
'	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeadingSmall><Td >" 
'	sSql = sSql & "Amount in RO '000"
'	sSql = sSql & "</Td >" 
'	sSql = sSql & "</TR>" 
'	sSql = sSql & "<TR Class=GreenTransparentHeadingVerySmall><Td>" 
'	sSql = sSql & "Company-wise Performance Summary of Group Results for the period:&nbsp;" & FullMonthsArr(cint(TheReportMonth)-1) & "/" & TheReportYear & "&nbsp;(YTD)"
'	sSql = sSql & "</Td >" 
'	sSql = sSql & "</TR>" 
'	sSql = sSql & "</Table>"
'	response.write sSql	


response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR >"
response.write "<TD valign=top align=left>"


call Page15("Y")

'response.write "<TR>"
'response.write "<TD align=center>"

'response.write "</TD>"
'response.write "</TR>"

response.write "</Table>"
response.write "<font class='pageNumber'>Page 16<font>"    
response.write "</div>"
response.write "</div>"
response.write "</div>"



response.write "<div class='page landscape-parent'>"
response.write "<div class='landscape'>"
response.write "<div class='content'>"

'	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
'	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
'	sSql = sSql & "Amount in RO '000"
'	sSql = sSql & "</Td >" 
'	sSql = sSql & "</TR>" 
'	sSql = sSql & "<TR Class=GreenTransparentHeading><Td >" 
'	sSql = sSql & "Sectorwise Summary of Group Results" 
'	sSql = sSql & "</Td >" 
'	sSql = sSql & "</TR>" 
'	sSql = sSql & "</Table>"
'	response.write sSql	

response.write "<Table border=0 width='99%' height='95%'>"
response.write "<TR >"
response.write "<TD valign=top>"


	call Page17()

'response.write "<TR>"
'response.write "<TD align=center>"

'response.write "</TD>"
'response.write "</TR>"

response.write "</Table>"
response.write "<font class='pageNumber'>Page 17<font>"
response.write "</div>"
response.write "</div>"
response.write "</div>"


'response.write "<div class='page'>"
response.write "<div class='content'>"
'response.write "<Table border=0 cellPadding=1 cellspacing=0 height='95%'  width='100%' >"
'response.write "<TR >"
'response.write "<TD valign=top>"
	call Annexures
'response.write "</TD>"
'response.write "</TR>"
'response.write "</Table>"

response.write "</div>"
'response.write "</div>"
	
	%>




	</Center>
	</Form>
	<Input Type=Hidden Name="txtVersion" Value="<%=Request.form("txtVersion")%>">
	<%If Trim(Request.Form("optPrint")) = "P" Then%>
		<script Language="javascript">
		window.document.execCommand("Print");
		window.close();
		</script>
	<%End if%>
	<%If Trim(Request.Form("optPrint")) = "S" Then%>
			<script Language="javascript">
			window.document.execCommand("SaveAs");
			window.close();
		</script>
	<%End if%>
	</Body>
</Html>

<%
	Private function printReport(sSql,tblHeading)

'response.write sSql

		rRs.Open Trim(sSql),cCn
		Content=""
		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""
		
		SalesCol1=""
		SalesCol2=""
		SalesCol3=""
		SalesCol4=""
		SalesCol5=""
		SalesCol6=""
		SalesCol7=""
		SalesCol8=""
		SalesCol9=""
		SalesCol10=""


		CommisionIncomeCol1=""
		CommisionIncomeCol2=""
		CommisionIncomeCol3=""
		CommisionIncomeCol4=""
		CommisionIncomeCol5=""
		CommisionIncomeCol6=""
		CommisionIncomeCol7=""
		CommisionIncomeCol8=""
		CommisionIncomeCol9=""
		CommisionIncomeCol10=""

		InvestmentIncomeCol1=""
		InvestmentIncomeCol2=""
		InvestmentIncomeCol3=""
		InvestmentIncomeCol4=""
		InvestmentIncomeCol5=""
		InvestmentIncomeCol6=""
		InvestmentIncomeCol7=""
		InvestmentIncomeCol8=""
		InvestmentIncomeCol9=""
		InvestmentIncomeCol10=""


		OtherIncome_LossCol1=""
		OtherIncome_LossCol2=""
		OtherIncome_LossCol3=""
		OtherIncome_LossCol4=""
		OtherIncome_LossCol5=""
		OtherIncome_LossCol6=""
		OtherIncome_LossCol7=""
		OtherIncome_LossCol8=""
		OtherIncome_LossCol9=""
		OtherIncome_LossCol10=""

		PBTCol1=""
		PBTCol2=""
		PBTCol3=""
		PBTCol4=""
		PBTCol5=""
		PBTCol6=""
		PBTCol7=""
		PBTCol8=""
		PBTCol9=""
		PBTCol10=""

		CashProfitLossCol1=""
		CashProfitLossCol2=""
		CashProfitLossCol3=""
		CashProfitLossCol4=""
		CashProfitLossCol5=""
		CashProfitLossCol6=""
		CashProfitLossCol7=""
		CashProfitLossCol8=""
		CashProfitLossCol9=""
		CashProfitLossCol10=""

		PBTFromOperationCol1=""
		PBTFromOperationCol2=""
		PBTFromOperationCol3=""
		PBTFromOperationCol4=""
		PBTFromOperationCol5=""
		PBTFromOperationCol6=""
		PBTFromOperationCol7=""
		PBTFromOperationCol8=""
		PBTFromOperationCol9=""
		PBTFromOperationCol10=""

		TotalRevenueCol1=""
		TotalRevenueCol2=""
		TotalRevenueCol3=""
		TotalRevenueCol4=""
		TotalRevenueCol5=""
		TotalRevenueCol6=""
		TotalRevenueCol7=""
		TotalRevenueCol8=""
		TotalRevenueCol9=""
		TotalRevenueCol10=""

		PBTPercentageTotalRevenueCol1=""
		PBTPercentageTotalRevenueCol2=""
		PBTPercentageTotalRevenueCol3=""
		PBTPercentageTotalRevenueCol4=""
		PBTPercentageTotalRevenueCol5=""
		PBTPercentageTotalRevenueCol6=""
		PBTPercentageTotalRevenueCol7=""
		PBTPercentageTotalRevenueCol8=""
		PBTPercentageTotalRevenueCol9=""
		PBTPercentageTotalRevenueCol10=""


		CashPBTPercentageTotalRevenueCol1=""
		CashPBTPercentageTotalRevenueCol2=""
		CashPBTPercentageTotalRevenueCol3=""
		CashPBTPercentageTotalRevenueCol4=""
		CashPBTPercentageTotalRevenueCol5=""
		CashPBTPercentageTotalRevenueCol6=""
		CashPBTPercentageTotalRevenueCol7=""
		CashPBTPercentageTotalRevenueCol8=""
		CashPBTPercentageTotalRevenueCol9=""
		CashPBTPercentageTotalRevenueCol10=""

		CashPBTMinusOperationCol1=""
		CashPBTMinusOperationCol2=""
		CashPBTMinusOperationCol3=""
		CashPBTMinusOperationCol4=""
		CashPBTMinusOperationCol5=""
		CashPBTMinusOperationCol6=""
		CashPBTMinusOperationCol7=""
		CashPBTMinusOperationCol8=""
		CashPBTMinusOperationCol9=""
		CashPBTMinusOperationCol10=""



		If Not rRs.Eof Then
			'
			sSql = ""

			sSql = sSql & "<Table border=0  width='100%' cellspacing=0 cellpadding=3 cellspacing=0 >"
			
			sSql = sSql & "<TR Class=AllReportHeader><Td style='padding: 0 2px 0 2px;' Align='center' colspan=10 class='border-ALL'>" & tblHeading & "</td></TR>"	
			sSql = sSql & "<TR Class=RepfontStyle><Td style='padding: 0 2px 0 2px;'  Align='center' width='30%' class='no-top-right-border'>Description</td>"	
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  Align='center' class='no-top-right-border'>Flash</td>"	
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  Align='center' class='no-top-right-border'>Budget</td>"	
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  Align='center' colspan=2 class='no-top-right-border'>vs Budget</td>"	
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  Align='center' class='no-top-right-border'>Variance<br>% vs<br>Budget</td>"	
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  Align='center' class='no-top-right-border'>Last Year<br>(LY)</td>"	
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  Align='center' colspan=2 class='no-top-right-border'>vs LY</td>"	
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  Align='center' class='no-top-border'>Variance<br>% vs<br>LY</td>"	
			sSql = sSql & "</tr>"

			Do while not rRs.EOF
				'	
				If InStr(trim(rRs("Description")), "%") then
					cDec=2
				Else
					cDec=1
				End if



		
				If rRs("DESCCD") ="A5" Then
					SalesCol1 = "<Td style='padding: 0 2px 0 2px;'  Align=Left  nowrap class='no-top-right-border_GREY'>" & rRs("Description")& "</Td>"
					'SalesCol2 = "<Td Align=right nowrap>" & formatnumber((rRs("MTD")),0,,(-1)) & "</Td>"
					SalesCol2 = rRs("MTD")
					SalesCol3 = rRs("MBDG")


					If rRs("VAR1") < 0 Then
						SalesCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR1") > 0 Then
						SalesCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						SalesCol4 = "<Td Align=center border=0  class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 


					SalesCol5 = rRs("VAR1") 
					'SalesCol6 = rRs("VARP") 

					If rRs("MBDG") <> 0 then
						SalesCol6 = (rRs("VAR1")/rRs("MBDG"))*100
					Else
						SalesCol6 = 0
					End If

					SalesCol7 = rRs("LY")
					If rRs("VAR") < 0 Then
						SalesCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR") > 0 Then
						SalesCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						SalesCol8 = "<Td Align=center border=0  class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 
					SalesCol9 = rRs("VAR") 
					'SalesCol10 = rRs("VARLY")
					If rRs("LY") <> 0 then
						SalesCol10 = (rRs("VAR")/rRs("LY"))*100
					Else
						SalesCol10 = 0
					End If

				End if 


		
				 If rRs("DESCCD") ="D5" Then
					'CommisionIncomeCol1 = "<Td style='padding: 0 2px 0 2px;'  Align=Left  nowrap>" & rRs("Description")& "</Td>"
					CommisionIncomeCol1 = "<Td style='padding: 0 2px 0 2px;'  Align=Left  nowrap class='no-top-right-border_Grey'>Commission Income</Td>"
					CommisionIncomeCol2 = rRs("MTD")
					CommisionIncomeCol3 = rRs("MBDG")

					If rRs("VAR1") < 0 Then
						CommisionIncomeCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR1") > 0 Then
						CommisionIncomeCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						CommisionIncomeCol4 = "<Td Align=center border=0 class='no-top-right-border_Grey'><img src='images/zero.png' width='80%'></Td>"
					End If 



					CommisionIncomeCol5 = rRs("VAR1")
					'CommisionIncomeCol6 = rRs("VARP")
					If rRs("MBDG") <> 0 then
						CommisionIncomeCol6 = (rRs("VAR1")/rRs("MBDG"))*100
					Else
						CommisionIncomeCol6 = 0
					End If
					
					CommisionIncomeCol7 = rRs("LY")

					If rRs("VAR") < 0 Then
						CommisionIncomeCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png'></Td>"
					ElseIf rRs("VAR") > 0 Then
						CommisionIncomeCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						CommisionIncomeCol8 = "<Td Align=center border=0 class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 

				
					CommisionIncomeCol9 = rRs("VAR")
					'CommisionIncomeCol10 = rRs("VARLY")
					If rRs("LY") <> 0 then
						CommisionIncomeCol10 = (rRs("VAR")/rRs("LY"))*100
					Else
						CommisionIncomeCol10 = 0
					End If

				End if 

				If rRs("DESCCD") ="E6" Then
					'OtherIncome_LossCol1 = "<Td Align=Left  nowrap>" & rRs("Description")& "</Td>"
					OtherIncome_LossCol1 = "<Td style='padding: 0 2px 0 2px;'  Align=Left  nowrap class='no-top-right-border_Grey'>Other Income</Td>"
					OtherIncome_LossCol2 = rRs("MTD")
					OtherIncome_LossCol3 = rRs("MBDG")

					If rRs("VAR1") < 0 Then
						OtherIncome_LossCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png'></Td>"
					ElseIf rRs("VAR1") > 0 Then
						OtherIncome_LossCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
					Else
						OtherIncome_LossCol4 = "<Td Align=center border=0  class='no-top-right-border'><img src='images/zero.png'  width='80%' ></Td>"
					End If 


					OtherIncome_LossCol5 =  rRs("VAR1")
					'OtherIncome_LossCol6 = rRs("VARP")

					If rRs("MBDG") <> 0 then
						OtherIncome_LossCol6 = (rRs("VAR1")/rRs("MBDG"))*100
					Else
						OtherIncome_LossCol6 = 0
					End If

					OtherIncome_LossCol7 = rRs("LY")

					If rRs("VAR") < 0 Then
						OtherIncome_LossCol8 = "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png'></Td>"
					ElseIf rRs("VAR") > 0 Then
						OtherIncome_LossCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png'></Td>"
					Else
						OtherIncome_LossCol8 = "<Td Align=center border=0 class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
					End If 

					OtherIncome_LossCol9 = rRs("VAR")
					'OtherIncome_LossCol10 = rRs("VARLY")


					If rRs("LY") <> 0 then
						OtherIncome_LossCol10 = (rRs("VAR")/rRs("LY"))*100
					Else
						OtherIncome_LossCol10 = 0
					End If

				End if 

				If rRs("DESCCD") ="E5" Then
					InvestmentIncomeCol1 = "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' Align=Left  nowrap>" & rRs("Description")& "</Td>"
					InvestmentIncomeCol2 = rRs("MTD")
					InvestmentIncomeCol3 = rRs("MBDG")

					If rRs("VAR1") < 0 Then
						InvestmentIncomeCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR1") > 0 Then
						InvestmentIncomeCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						InvestmentIncomeCol4 = "<Td Align=center border=0 class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 


					InvestmentIncomeCol5 = rRs("VAR1")
					'InvestmentIncomeCol6 = rRs("VARP")


					If rRs("MBDG") <> 0 then
						InvestmentIncomeCol6 = (rRs("VAR1")/rRs("MBDG"))*100
					Else
						InvestmentIncomeCol6 = 0
					End If

					InvestmentIncomeCol7 = rRs("LY")

					If rRs("VAR") < 0 Then
						InvestmentIncomeCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR") > 0 Then
						InvestmentIncomeCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						InvestmentIncomeCol8 = "<Td Align=center border=0  class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 

					InvestmentIncomeCol9 = rRs("VAR")
					'InvestmentIncomeCol10 = rRs("VARLY")

					If rRs("LY") <> 0 then
						InvestmentIncomeCol10 = (rRs("VAR")/rRs("LY"))*100
					Else
						InvestmentIncomeCol10 = 0
					End If
					
				End if 

				If rRs("DESCCD") ="R5" Then
					PBTFromOperationCol1 = "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' Align=Left  nowrap>PBT-Operations</Td>"
					PBTFromOperationCol2 = rRs("MTD")
					PBTFromOperationCol3 = rRs("MBDG")

					If rRs("VAR1") < 0 Then
						PBTFromOperationCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR1") > 0 Then
						PBTFromOperationCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						PBTFromOperationCol4 = "<Td Align=center border=0  class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 



					PBTFromOperationCol5 = rRs("VAR1")
					'PBTFromOperationCol6 = rRs("VARP")
					
					If rRs("MBDG") <> 0 then
						PBTFromOperationCol6 = (rRs("VAR1")/rRs("MBDG"))*100
					Else
						PBTFromOperationCol6 = 0
					End If


					PBTFromOperationCol7 = rRs("LY")

					If rRs("VAR") < 0 Then
						PBTFromOperationCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR") > 0 Then
						PBTFromOperationCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						PBTFromOperationCol8 = "<Td Align=center border=0  class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 

					
					PBTFromOperationCol9 = rRs("VAR")
					'PBTFromOperationCol10 = rRs("VARLY")
					If rRs("LY") <> 0 then
						PBTFromOperationCol10 = (rRs("VAR")/rRs("LY"))*100
					Else
						PBTFromOperationCol10 = 0
					End If

				End if 

				If rRs("DESCCD") ="E5" Then
					InvestmentIncomeCol1 = "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' Align=Left  nowrap>Add: Investment Income</Td>"
					InvestmentIncomeCol2 = rRs("MTD")
					InvestmentIncomeCol3 = rRs("MBDG")

					If rRs("VAR1") < 0 Then
						InvestmentIncomeCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR1") > 0 Then
						InvestmentIncomeCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						InvestmentIncomeCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 


					InvestmentIncomeCol5 = rRs("VAR1")
					'InvestmentIncomeCol6 = rRs("VARP")

					If rRs("MBDG") <> 0 then
						InvestmentIncomeCol6 = (rRs("VAR1")/rRs("MBDG"))*100 
					Else
						InvestmentIncomeCol6 = 0
					End If

					InvestmentIncomeCol7 = rRs("LY")
					If rRs("VAR") < 0 Then
						InvestmentIncomeCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR") > 0 Then
						InvestmentIncomeCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						InvestmentIncomeCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 

					InvestmentIncomeCol9 = rRs("VAR")
					'InvestmentIncomeCol10 = rRs("VARLY")
					If rRs("LY") <> 0 then
						InvestmentIncomeCol10 = (rRs("VAR")/rRs("LY"))*100
					Else
						InvestmentIncomeCol10 = 0
					End If

				End if 

				If rRs("DESCCD") ="K5" Then
					'PBTCol1 = "<Td Align=Left  nowrap>" & rRs("Description")& "</Td>"
					PBTCol1 = "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' Align=Left  nowrap>PBT-Company</Td>"
					PBTCol2 = rRs("MTD")
					PBTCol3 = rRs("MBDG")

					If rRs("VAR1") < 0 Then
						PBTCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR1") > 0 Then
						PBTCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						PBTCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 


					PBTCol5 = rRs("VAR1")
					'PBTCol6 = rRs("VARP")

					If rRs("MBDG") <> 0 then
						PBTCol6 = (rRs("VAR1")/rRs("MBDG"))*100
					Else
						PBTCol6 = 0
					End If

					PBTCol7 = rRs("LY")

					If rRs("VAR") < 0 Then
						PBTCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR") > 0 Then
						PBTCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
					Else
						PBTCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png'  width='80%'></Td>"
					End If 


					PBTCol9 = rRs("VAR")
					'PBTCol10 = rRs("VARLY")


					If rRs("LY") <> 0 then
						PBTCol10 = (rRs("VAR")/rRs("LY"))*100
					Else
						PBTCol10 = 0
					End If

				End if 

				If rRs("DESCCD") ="M5" Then
					'CashProfitLossCol1 = "<Td Align=Left  nowrap>" & rRs("Description") & "</Td>"
					CashProfitLossCol1 = "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' Align=Left  nowrap>Cash PBT - Company</Td>"
					CashProfitLossCol2 = rRs("MTD")
					CashProfitLossCol3 = rRs("MBDG")

					If rRs("VAR1") < 0 Then
						CashProfitLossCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR1") > 0 Then
						CashProfitLossCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
					Else
						CashProfitLossCol4 = "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
					End If 


					CashProfitLossCol5 = rRs("VAR1")
					'CashProfitLossCol6 = rRs("VARP")


					If rRs("MBDG") <> 0 then
						CashProfitLossCol6 = (rRs("VAR1")/rRs("MBDG"))*100
					Else
						CashProfitLossCol6 = 0
					End If


					CashProfitLossCol7 = rRs("LY")

					If rRs("VAR") < 0 Then
						CashProfitLossCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
					ElseIf rRs("VAR") > 0 Then
						CashProfitLossCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
					Else
						CashProfitLossCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
					End If 

				
					CashProfitLossCol9 = rRs("VAR")
					'CashProfitLossCol10 = rRs("VARLY")
					If rRs("LY") <> 0 then
						CashProfitLossCol10 = (rRs("VAR")/rRs("LY"))*100
					Else
						CashProfitLossCol10 = 0
					End If

				End if 



	
				rRs.MoveNext
				
			Loop
			
			'RepContentFont
			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & SalesCol1
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If SalesCol2 <> 0 then
				sSql = sSql & formatnumber(SalesCol2,0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If SalesCol3 <> 0 then
				sSql = sSql & formatnumber((SalesCol3),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 

			sSql = sSql & "</TD>"
			sSql = sSql & SalesCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If SalesCol5 <> 0 then
				sSql = sSql & formatnumber((SalesCol5),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If SalesCol6 <> 0 THEN
				sSql = sSql & formatnumber((SalesCol6),1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If 
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			
			If SalesCol7 <> 0 then
				sSql = sSql & formatnumber((SalesCol7),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 

			sSql = sSql & "</TD>"
			sSql = sSql & SalesCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"

			If SalesCol9 <> 0 then
				sSql = sSql & formatnumber((SalesCol9),0,,(-1))  
			Else
				sSql = sSql & "&nbsp;"
			End If 

			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border_Grey' align=right>"
			If SalesCol10 <> 0 then
				sSql = sSql & formatnumber((SalesCol10),1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & CommisionIncomeCol1
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If CommisionIncomeCol2 <> 0 then
				sSql = sSql & formatnumber((CommisionIncomeCol2),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If CommisionIncomeCol3 <> 0 then
				sSql = sSql & formatnumber((CommisionIncomeCol3),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			sSql = sSql & "</TD>"
			sSql = sSql & CommisionIncomeCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If CommisionIncomeCol5 <>0 then
				sSql = sSql & formatnumber((CommisionIncomeCol5),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If CommisionIncomeCol6 <> 0 THEN
				sSql = sSql & formatnumber((CommisionIncomeCol6),1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If 


			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If CommisionIncomeCol7 <> 0 then
				sSql = sSql & formatnumber((CommisionIncomeCol7),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 
			sSql = sSql & "</TD>"
			sSql = sSql & CommisionIncomeCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If CommisionIncomeCol9 <> 0 then
				sSql = sSql & formatnumber((CommisionIncomeCol9),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border_Grey' align=right>"
			If CommisionIncomeCol10 <> 0 then
				sSql = sSql & formatnumber((CommisionIncomeCol10),1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End if

			

			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & OtherIncome_LossCol1
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			If OtherIncome_LossCol2 <> 0 then
				sSql = sSql & formatnumber((OtherIncome_LossCol2),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if

			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			If OtherIncome_LossCol3 <> 0 then
				sSql = sSql & formatnumber((OtherIncome_LossCol3),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if


			sSql = sSql & "</TD>"
			sSql = sSql & OtherIncome_LossCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border' align=right>"
			If OtherIncome_LossCol5 <> 0 Then   ' Vipul - Do the same for other columns 31-Mar-2020
				sSql = sSql & formatnumber((OtherIncome_LossCol5),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 
			sSql = sSql & "</TD>"

			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			If OtherIncome_LossCol6 <> 0 THEN
				sSql = sSql & formatnumber((OtherIncome_LossCol6),1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If 


			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			If OtherIncome_LossCol7 <> 0 then
				sSql = sSql & formatnumber((OtherIncome_LossCol7),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 

			
			sSql = sSql & "</TD>"
			sSql = sSql & OtherIncome_LossCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border' align=right>"
			If OtherIncome_LossCol9 <> 0 Then '''''Becareful with zero
				sSql = sSql & formatnumber((OtherIncome_LossCol9),0,,(-1)) 
			Else 
				sSql = sSql & "&nbsp;"
			End If 
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border' align=right>"
			If OtherIncome_LossCol10 <> 0 then
				sSql = sSql & formatnumber((OtherIncome_LossCol10),1,,(-1)) & "%" 
			Else
				sSql = sSql & "&nbsp;"
			End if


			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"


'response.write "<br>" &	SalesCol2 
'response.write "<br>" & CommisionIncomeCol2 
'response.write "<br>" & OtherIncome_LossCol2

			' "Total Revenue" printing start
			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' >"
			sSql = sSql & "Total Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			
			If ((SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2) <> 0) then
				sSql = sSql & formatnumber((SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 

			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			If ((SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3) <> 0) then
				sSql = sSql & formatnumber((SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3),0,,(-1))  
			Else
				sSql = sSql & "&nbsp;"
			End If 
			
			sSql = sSql & "</TD>"


			If (SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3) < 0 Then
				TotalRevenueCol4 = "<Td style='padding: 0 2px 0 2px;'  Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3) > 0 then
				TotalRevenueCol4 = "<Td style='padding: 0 2px 0 2px;'  Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			Else
				TotalRevenueCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
			End If 

		
			sSql = sSql & TotalRevenueCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border' align=right>"
			If ((SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3) <> 0) then
				sSql = sSql & formatnumber((SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End If 
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((((SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100,1,,(-1))  & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7),0,,(-1)) 
			sSql = sSql & "</TD>"

			If (SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7) < 0 Then
				TotalRevenueCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7) > 0 Then
				TotalRevenueCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			Else
				TotalRevenueCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
			End If 

			
			sSql = sSql & TotalRevenueCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border' align=right>"
			sSql = sSql & formatnumber((SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border' align=right>"
			sSql = sSql & formatnumber((((SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)-(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))/(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"


			' "Total Revenue" printing End




			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & PBTFromOperationCol1
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If PBTFromOperationCol2 <> 0 then
				sSql = sSql & formatnumber((PBTFromOperationCol2),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if

			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If PBTFromOperationCol3 <> 0 then
				sSql = sSql & formatnumber((PBTFromOperationCol3),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if

			sSql = sSql & "</TD>"
			sSql = sSql & PBTFromOperationCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If PBTFromOperationCol5 <> 0 then
				sSql = sSql & formatnumber((PBTFromOperationCol5),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If PBTFromOperationCol6 <> 0 Then
				If cdbl(PBTFromOperationCol6) >= 1000 Then
					sSql = sSql & ">1000%"
				ElseIf CDbl(PBTFromOperationCol6) <= -1000 Then
					sSql = sSql & ">(1000%)" 
				Else
					sSql = sSql & formatnumber((PBTFromOperationCol6),1,,(-1)) & "%"
				End if
				'
			Else
				'sSql = sSql & "&nbsp;*****"
				sSql = sSql & formatnumber((PBTFromOperationCol6),1,,(-1)) & "%"
			End If 


			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If PBTFromOperationCol7 <> 0 then
				sSql = sSql & formatnumber((PBTFromOperationCol7),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if

			
			sSql = sSql & "</TD>"
			sSql = sSql & PBTFromOperationCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If PBTFromOperationCol9 <> 0 then
				sSql = sSql & formatnumber((PBTFromOperationCol9),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border_Grey' align=right>"

			
			'If PBTFromOperationCol10 <> 0 then
			'	sSql = sSql & formatnumber((PBTFromOperationCol10),1,,(-1)) & "%" 
			'Else
			'	sSql = sSql & "&nbsp;"
			'End if

			If PBTFromOperationCol10 <> 0 Then
				If cdbl(PBTFromOperationCol10) >= 1000 Then
					sSql = sSql & ">1000%"
				ElseIf CDbl(PBTFromOperationCol10) <= -1000 Then
					sSql = sSql & ">(1000%)" 
				Else
					sSql = sSql & formatnumber((PBTFromOperationCol10),1,,(-1)) & "%"
				End if
				'
			Else
				sSql = sSql & "&nbsp;"
			End If 



			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"



		' "PBT % on Total Revenue" printing start

			sSql = sSql & "<Tr Class=RepfontStyleItalicPrintReport>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' >"
			sSql = sSql & "PBT % on Total Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol2/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol3/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"


			If ((PBTFromOperationCol2/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)*100)-(PBTFromOperationCol3/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100) < 0 Then
				TotalRevenueCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf ((PBTFromOperationCol2/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)*100)-(PBTFromOperationCol3/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100) > 0 Then
				TotalRevenueCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				TotalRevenueCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
			End If 

		
			sSql = sSql & TotalRevenueCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			sSql = sSql &  formatnumber(((PBTFromOperationCol2/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2)*100)-(PBTFromOperationCol3/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100),1,,(-1))  & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			sSql = sSql & "&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol7 / (SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			If ((PBTFromOperationCol2/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100)-((PBTFromOperationCol7 / (SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100) < 0 Then
				TotalRevenueCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf ((PBTFromOperationCol2/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100)-((PBTFromOperationCol7 / (SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100) > 0 Then
				TotalRevenueCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				TotalRevenueCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
			End If 


			
			sSql = sSql & TotalRevenueCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			sSql = sSql & formatnumber(((PBTFromOperationCol2/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100)-((PBTFromOperationCol7 / (SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100),1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border_Grey' align=right>"
			sSql = sSql & "&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			' "PBT % on Total Revenue" printing End


			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & InvestmentIncomeCol1
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If InvestmentIncomeCol2 <> 0 then
				sSql = sSql & formatnumber((InvestmentIncomeCol2),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if

			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If InvestmentIncomeCol3 <> 0 then
				sSql = sSql & formatnumber((InvestmentIncomeCol3),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if

			sSql = sSql & "</TD>"
			sSql = sSql & InvestmentIncomeCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If InvestmentIncomeCol5 <> 0 then
				sSql = sSql & formatnumber((InvestmentIncomeCol5),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;" 
			End if
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If InvestmentIncomeCol6 <> 0 THEN
				'sSql = sSql & formatnumber((InvestmentIncomeCol6),1,,(-1)) & "%"




				
					If cdbl(InvestmentIncomeCol6) >= 1000 Then
						sSql = sSql & ">1000%"
					ElseIf CDbl(InvestmentIncomeCol6) <= -1000 Then
						sSql = sSql & ">(1000%)" 
					Else
						sSql = sSql & formatnumber((InvestmentIncomeCol6),1,,(-1)) & "%"
					End if
					'
				

			Else
				sSql = sSql & "&nbsp;"
			End If 


			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If InvestmentIncomeCol7 <> 0 then
				sSql = sSql & formatnumber((InvestmentIncomeCol7),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			sSql = sSql & "</TD>"
			sSql = sSql & InvestmentIncomeCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If InvestmentIncomeCol9 <> 0 then
				sSql = sSql & formatnumber((InvestmentIncomeCol9),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border_Grey' align=right>"

			
			' Following line was added on 31-Jan-2023
			If cdbl(InvestmentIncomeCol10) >= 1000 Then
				sSql = sSql & ">1000%"
			ElseIf CDbl(InvestmentIncomeCol10) <= -1000 Then
				sSql = sSql & ">(1000%)" 
			ElseIf InvestmentIncomeCol10 = 0 Then
				sSql = sSql & "&nbsp;" 			
			Else
				sSql = sSql & formatnumber((InvestmentIncomeCol10),1,,(-1)) & "%"
			End If
					
			'If InvestmentIncomeCol10 <> 0 then
			'	sSql = sSql & formatnumber((InvestmentIncomeCol10),1,,(-1)) & "%" 
			'Else
			'	sSql = sSql & "&nbsp;"
			'End if


			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & PBTCol1
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If PBTCol2 <> 0 then
				sSql = sSql & formatnumber((PBTCol2),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If PBTCol3 <> 0 then
				sSql = sSql & formatnumber((PBTCol3),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & PBTCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If PBTCol5 <> 0 then
				sSql = sSql & formatnumber((PBTCol5),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If PBTCol6 <> 0 THEN
				'sSql = sSql & formatnumber((PBTCol6),1,,(-1)) & "%"

					' 28-Feb-2021
					If cdbl(PBTCol6) >= 1000 Then
						sSql = sSql & ">1000%"
					ElseIf CDbl(PBTCol6) <= -1000 Then
						sSql = sSql & ">(1000%)" 
					Else
						sSql = sSql & formatnumber((PBTCol6),1,,(-1)) & "%"
					End If
					
			Else
				sSql = sSql & "&nbsp;"
			End If 


			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			If PBTCol7 <> 0 then
				sSql = sSql & formatnumber((PBTCol7),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & PBTCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			If PBTCol9 <> 0 then
				sSql = sSql & formatnumber((PBTCol9),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border_Grey' align=right>"

			If PBTCol10 <> 0 Then
				'sSql = sSql & formatnumber((PBTCol10),1,,(-1)) & "%"  ' following changed on 28-Jul-2022
				If (PBTCol10  >= 1000) Then
					sSql = sSql & ">1000%"
				ElseIf (PBTCol10 <= -1000) Then
					sSql = sSql & ">(1000%)" 
				Else
					sSql = sSql & formatnumber((PBTCol10),1,,(-1)) & "%"
				End If
			Else
				sSql = sSql & "&nbsp;"
			End if

			
				
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"


			' "Cash PBT - Operations" printing start

			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' bgcolor='#D8E4BC'>"
			sSql = sSql & "Cash PBT - Operations"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol2-InvestmentIncomeCol2,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol3-InvestmentIncomeCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			If ((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol3-InvestmentIncomeCol3)) < 0 Then
				TotalRevenueCol4 = "<TD class='no-top-right-border_Grey' Align=center border=0 width='30' bgcolor='#D8E4BC'><img src='images/down.png'></Td>"
			ElseIf ((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol3-InvestmentIncomeCol3)) > 0 Then
				TotalRevenueCol4 = "<TD class='no-top-right-border_Grey' Align=center border=0 width='30' bgcolor='#D8E4BC'><img src='images/up.png'></Td>"
			Else
				TotalRevenueCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border_Grey' bgcolor='#D8E4BC'><img src='images/zero.png'  width='80%'></Td>"
			End If 

			sSql = sSql & TotalRevenueCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol3-InvestmentIncomeCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"

			' The following if condition was included on 27-May-2021
				If ((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol3-InvestmentIncomeCol3))/(CashProfitLossCol3-InvestmentIncomeCol3)*100 >= 1000 Then
					sSql = sSql & ">1000%"
				ElseIf ((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol3-InvestmentIncomeCol3))/(CashProfitLossCol3-InvestmentIncomeCol3)*100 <= -1000 Then
					sSql = sSql & ">(1000%)" 
				Else
					sSql = sSql & formatnumber(((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol3-InvestmentIncomeCol3))/(CashProfitLossCol3-InvestmentIncomeCol3)*100,1,,(-1))  & "%"
				End if

			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol7-InvestmentIncomeCol7,0,,(-1)) 
			sSql = sSql & "</TD>"

			If ( (CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol7-InvestmentIncomeCol7)  ) < 0 Then
				TotalRevenueCol8 = "<TD class='no-top-right-border_Grey' Align=center border=0 width='30' bgcolor='#D8E4BC'><img src='images/down.png' ></Td>"
			ElseIf ( (CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol7-InvestmentIncomeCol7)  ) > 0 Then
				TotalRevenueCol8 = "<TD class='no-top-right-border_Grey' Align=center border=0 width='30' bgcolor='#D8E4BC'><img src='images/up.png' ></Td>"
			Else
				TotalRevenueCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border' bgcolor='#D8E4BC'><img src='images/zero.png'  width='80%'></Td>"
			End If 


			sSql = sSql & TotalRevenueCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol7-InvestmentIncomeCol7),0,,(-1))
			sSql = sSql & "</TD>"



			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border_Grey' align=right bgcolor='#D8E4BC'>"



			If cdbl(  (((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol7-InvestmentIncomeCol7))/(CashProfitLossCol7-InvestmentIncomeCol7)) * 100 ) >= 1000 Then
				sSql = sSql & ">1000%"
			ElseIf cdbl(  (((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol7-InvestmentIncomeCol7))/(CashProfitLossCol7-InvestmentIncomeCol7)) * 100 ) <= -1000 Then
				sSql = sSql & ">(1000%)" 
			Else
				sSql = sSql & formatnumber((((CashProfitLossCol2-InvestmentIncomeCol2)-(CashProfitLossCol7-InvestmentIncomeCol7))/(CashProfitLossCol7-InvestmentIncomeCol7))*100,1,,(-1)) & "%"
			End if

			
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			' "Cash PBT - Operations" printing End


	' "Cash PBT % on Total Revenue" printing start

			sSql = sSql & "<Tr Class=RepfontStyleItalicPrintReport>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey'>"
			sSql = sSql & "Cash PBT % on Total Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol3-InvestmentIncomeCol3)/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

'((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100-((CashProfitLossCol3-InvestmentIncomeCol3)/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100

			If (((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100-((CashProfitLossCol3-InvestmentIncomeCol3)/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100) < 0 Then
				TotalRevenueCol4 = "<TD class='no-top-right-border_Grey' Align=center border=0 width='30' ><img src='images/down.png' ></Td>"
			ElseIf (((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100-((CashProfitLossCol3-InvestmentIncomeCol3)/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100) > 0 Then
				TotalRevenueCol4 = "<TD class='no-top-right-border_Grey' Align=center border=0 width='30' ><img src='images/up.png' ></Td>"
			Else
				TotalRevenueCol4 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
			End If 

'response.write ((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100-((CashProfitLossCol3-InvestmentIncomeCol3)/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100


			sSql = sSql & TotalRevenueCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"
			sSql = sSql &  formatnumber(((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100-((CashProfitLossCol3-InvestmentIncomeCol3)/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100,1,,(-1))  & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & "&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol7-InvestmentIncomeCol7)/(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			If ((((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100)-(((CashProfitLossCol7-InvestmentIncomeCol7)/(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100)) < 0 Then
				TotalRevenueCol8 = "<TD class='no-top-right-border_Grey' Align=center border=0 width='30' ><img src='images/down.png' ></Td>"
			ElseIf ((((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100)-(((CashProfitLossCol7-InvestmentIncomeCol7)/(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100)) > 0 Then
				TotalRevenueCol8 = "<TD class='no-top-right-border_Grey' Align=center border=0 width='30' ><img src='images/up.png' ></Td>"
			Else
				TotalRevenueCol8 = "<Td Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'  width='80%'></Td>"
			End If 



'(((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100)-(((CashProfitLossCol7-InvestmentIncomeCol7)/(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100)
			
			sSql = sSql & TotalRevenueCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border_Grey' align=right>"

			sSql = sSql & FormatNumber((((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100)-(((CashProfitLossCol7-InvestmentIncomeCol7)/(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100),1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right>"
			sSql = sSql & "&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			' "Cash PBT % on Total Revenue" printing End


			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & CashProfitLossCol1
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			If CashProfitLossCol2 <> 0 then
				sSql = sSql & formatnumber((CashProfitLossCol2),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if

			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			If CashProfitLossCol3 <> 0 then
				sSql = sSql & formatnumber((CashProfitLossCol3),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if

			
			sSql = sSql & "</TD>"
			sSql = sSql & CashProfitLossCol4
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border' align=right>"
			If CashProfitLossCol5 <> 0 then
				sSql = sSql & formatnumber((CashProfitLossCol5),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			If CashProfitLossCol6 <> 0 Then ' following changed on 28-Jul-2022
				'sSql = sSql & formatnumber((CashProfitLossCol6),1,,(-1)) & "%"

				If (CashProfitLossCol6  >= 1000) Then
					sSql = sSql & ">1000%"
				ElseIf (CashProfitLossCol6 <= -1000) Then
					sSql = sSql & ">(1000%)" 
				Else
					sSql = sSql & formatnumber((CashProfitLossCol6),1,,(-1)) & "%"
				End If

			Else
				sSql = sSql & "&nbsp;"
			End If 


			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-right-border' align=right>"
			
			If CashProfitLossCol7 <> 0 then
				sSql = sSql & formatnumber((CashProfitLossCol7),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			sSql = sSql & "</TD>"
			sSql = sSql & CashProfitLossCol8
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-right-left-top-border' align=right>"
			If CashProfitLossCol9 <> 0 then
				sSql = sSql & formatnumber((CashProfitLossCol9),0,,(-1)) 
			Else
				sSql = sSql & "&nbsp;"
			End if
			
			sSql = sSql & "</TD>"
			sSql = sSql & "<Td style='padding: 0 2px 0 2px;'  class='no-top-border' align=right>"
			If CashProfitLossCol10 <> 0 Then
			
				If cdbl(CashProfitLossCol10) >= 1000 Then
					sSql = sSql & ">1000%"
				ElseIf CDbl(CashProfitLossCol10) <= -1000 Then
					sSql = sSql & ">(1000%)" 
				Else
					sSql = sSql & formatnumber((CashProfitLossCol10),1,,(-1)) & "%"
				End if
				'


			Else
				sSql = sSql & "&nbsp;"
			End if


			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			sSql = sSql & "</Table>"

			'		
			Response.write sSql 

		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (A)**********</Td>"
		End if		
		rRs.close
	End Function


	Private function Summary5Years()

		sSql = " select"
		sSql = sSql & " DESCCD, DESCRIPTION, "
		sSql = sSql & " SUM(ACT_YTD_BAL) YTD, "
		sSql = sSql & " SUM(BGT_YTD_BAL) MBDG, "
		sSql = sSql & " SUM(ACT_YTD_BAL_py1) PY1, "
		sSql = sSql & " SUM(ACT_YTD_BAL_PY2) Py2, "
		sSql = sSql & " SUM(ACT_YTD_BAL_PY3) PY3, "
		sSql = sSql & " SUM(ACT_YTD_BAL_PY4) PY4 "
		sSql = sSql & " from mistxn "
		sSql = sSql & " where"
		sSql = sSql & " cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' AND "
		sSql = sSql & " CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
		sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5') "
		sSql = sSql & " GROUP BY CYEAR, CMONTH, "
		sSql = sSql & " DESCCD, DESCRIPTION"
		sSql = sSql & " ORDER BY 1 "


		rRs.Open Trim(sSql),cCn
		Content=""
		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""
		
		SalesCol1=""
		SalesCol2=""
		SalesCol3=""
		SalesCol4=""
		SalesCol5=""
		SalesCol6=""
		SalesCol7=""


		CommisionIncomeCol1=""
		CommisionIncomeCol2=""
		CommisionIncomeCol3=""
		CommisionIncomeCol4=""
		CommisionIncomeCol5=""
		CommisionIncomeCol6=""
		CommisionIncomeCol7=""

		InvestmentIncomeCol1=""
		InvestmentIncomeCol2=""
		InvestmentIncomeCol3=""
		InvestmentIncomeCol4=""
		InvestmentIncomeCol5=""
		InvestmentIncomeCol6=""
		InvestmentIncomeCol7=""


		OtherIncome_LossCol1=""
		OtherIncome_LossCol2=""
		OtherIncome_LossCol3=""
		OtherIncome_LossCol4=""
		OtherIncome_LossCol5=""
		OtherIncome_LossCol6=""
		OtherIncome_LossCol7=""

		PBTCol1=""
		PBTCol2=""
		PBTCol3=""
		PBTCol4=""
		PBTCol5=""
		PBTCol6=""
		PBTCol7=""

		CashProfitLossCol1=""
		CashProfitLossCol2=""
		CashProfitLossCol3=""
		CashProfitLossCol4=""
		CashProfitLossCol5=""
		CashProfitLossCol6=""
		CashProfitLossCol7=""

		PBTFromOperationCol1=""
		PBTFromOperationCol2=""
		PBTFromOperationCol3=""
		PBTFromOperationCol4=""
		PBTFromOperationCol5=""
		PBTFromOperationCol6=""
		PBTFromOperationCol7=""

		TotalRevenueCol1=""
		TotalRevenueCol2=""
		TotalRevenueCol3=""
		TotalRevenueCol4=""
		TotalRevenueCol5=""
		TotalRevenueCol6=""
		TotalRevenueCol7=""

		PBTPercentageTotalRevenueCol1=""
		PBTPercentageTotalRevenueCol2=""
		PBTPercentageTotalRevenueCol3=""
		PBTPercentageTotalRevenueCol4=""
		PBTPercentageTotalRevenueCol5=""
		PBTPercentageTotalRevenueCol6=""
		PBTPercentageTotalRevenueCol7=""


		CashPBTPercentageTotalRevenueCol1=""
		CashPBTPercentageTotalRevenueCol2=""
		CashPBTPercentageTotalRevenueCol3=""
		CashPBTPercentageTotalRevenueCol4=""
		CashPBTPercentageTotalRevenueCol5=""
		CashPBTPercentageTotalRevenueCol6=""
		CashPBTPercentageTotalRevenueCol7=""

		CashPBTMinusOperationCol1=""
		CashPBTMinusOperationCol2=""
		CashPBTMinusOperationCol3=""
		CashPBTMinusOperationCol4=""
		CashPBTMinusOperationCol5=""
		CashPBTMinusOperationCol6=""
		CashPBTMinusOperationCol7=""


		If Not rRs.Eof Then
			'
			sSql = ""

			sSql = "<Table border='0' width='100%' cellspacing=0 cellpadding=3>"
			
			sSql = sSql & "<TR Class=AllReportHeader><TD class='border-ALL' Align='center' colspan=7>Summary for the 5 years YTD</td></TR>"	

			sSql = sSql & "<TR Class=RepfontStyle><TD class='no-top-right-border' Align='center' width='30%'>Description</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash<br>YTD " & MonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Budget<br>YTD " & MonthsArr(cint(TheReportMonth)-1) & "&nbsp" & TheReportYear & "</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Actuals<br>YTD " & MonthsArr(cint(TheReportMonth)-1) & "&nbsp" & CInt(TheReportYear)-1 & "</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Actuals<br>YTD " & MonthsArr(cint(TheReportMonth)-1) & "&nbsp" & CInt(TheReportYear)-2 & "</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Actuals<br>YTD " & MonthsArr(cint(TheReportMonth)-1) & "&nbsp" & CInt(TheReportYear)-3 & "</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >Actuals<br>YTD " & MonthsArr(cint(TheReportMonth)-1) & "&nbsp" & CInt(TheReportYear)-4 & "</td>"	
			sSql = sSql & "</tr>"

			Do while not rRs.EOF
				'	
				If InStr(trim(rRs("Description")), "%") then
					cDec=2
				Else
					cDec=1
				End if



		
				If rRs("DESCCD") ="A5" Then
					SalesCol1 = "<TD class='no-top-right-border_Grey' Align=Left  nowrap>" & rRs("Description")& "</Td>"
					'SalesCol2 = "<TD class='no-top-right-border' Align=right nowrap>" & formatnumber((rRs("MTD")),0,,(-1)) & "</Td>"
					SalesCol2 = rRs("YTD")
					SalesCol3 = rRs("MBDG")
					SalesCol4 = rRs("PY1")
					SalesCol5 = rRs("PY2")
					SalesCol6 = rRs("PY3")
					SalesCol7 = rRs("PY4")
				End if 


		
				 If rRs("DESCCD") ="D5" Then
					'CommisionIncomeCol1 = "<TD class='no-top-right-border' Align=Left  nowrap>" & rRs("Description")& "</Td>"
					CommisionIncomeCol1 = "<TD class='no-top-right-border_Grey' Align=Left  nowrap>Commission Income</Td>"
					CommisionIncomeCol2 = rRs("YTD")
					CommisionIncomeCol3 = rRs("MBDG")
					CommisionIncomeCol4 = rRs("PY1")
					CommisionIncomeCol5 = rRs("PY2")
					CommisionIncomeCol6 = rRs("PY3")
					CommisionIncomeCol7 = rRs("PY4")

				End if 

				If rRs("DESCCD") ="E6" Then
					'OtherIncome_LossCol1 = "<TD class='no-top-right-border' Align=Left  nowrap>" & rRs("Description")& "</Td>"

					OtherIncome_LossCol1 = "<TD class='no-top-right-border_Grey' Align=Left  nowrap>Other Income</Td>"
					OtherIncome_LossCol2 = rRs("YTD")
					OtherIncome_LossCol3 = rRs("MBDG")
					OtherIncome_LossCol4 = rRs("PY1")
					OtherIncome_LossCol5 = rRs("PY2")
					OtherIncome_LossCol6 = rRs("PY3")
					OtherIncome_LossCol7 = rRs("PY4")


				End if 

				If rRs("DESCCD") ="E5" Then
					InvestmentIncomeCol1 = "<TD class='no-top-right-border_Grey' Align=Left  nowrap>" & rRs("Description")& "</Td>"
					InvestmentIncomeCol2 = rRs("YTD")
					InvestmentIncomeCol3 = rRs("MBDG")
					InvestmentIncomeCol4 = rRs("PY1")
					InvestmentIncomeCol5 = rRs("PY2")
					InvestmentIncomeCol6 = rRs("PY3")
					InvestmentIncomeCol7 = rRs("PY4")

				End if 

				If rRs("DESCCD") ="R5" Then
					PBTFromOperationCol1 = "<TD class='no-top-right-border_Grey' Align=Left  nowrap>" & rRs("Description")& "</Td>"
					PBTFromOperationCol2 = rRs("YTD")
					PBTFromOperationCol3 = rRs("MBDG")
					PBTFromOperationCol4 = rRs("PY1")
					PBTFromOperationCol5 = rRs("PY2")
					PBTFromOperationCol6 = rRs("PY3")
					PBTFromOperationCol7 = rRs("PY4")
				End if 

				If rRs("DESCCD") ="E5" Then
					InvestmentIncomeCol1 = "<TD class='no-top-right-border_Grey' Align=Left  nowrap>" & rRs("Description")& "</Td>"
					InvestmentIncomeCol2 = rRs("YTD")
					InvestmentIncomeCol3 = rRs("MBDG")
					InvestmentIncomeCol4 = rRs("PY1")
					InvestmentIncomeCol5 = rRs("PY2")
					InvestmentIncomeCol6 = rRs("PY3")
					InvestmentIncomeCol7 = rRs("PY4")
				End if 

				If rRs("DESCCD") ="K5" Then
					'PBTCol1 = "<TD class='no-top-right-border' Align=Left  nowrap>" & rRs("Description")& "</Td>"
					PBTCol1 = "<TD class='no-top-right-border_Grey' Align=Left  nowrap>PBT-Company</Td>"
					PBTCol2 = rRs("YTD")
					PBTCol3 = rRs("MBDG")
					PBTCol4 = rRs("PY1")
					PBTCol5 = rRs("PY2")
					PBTCol6 = rRs("PY3")
					PBTCol7 = rRs("PY4")
				End if 

				If rRs("DESCCD") ="M5" Then
					'CashProfitLossCol1 = "<TD class='no-top-right-border' Align=Left  nowrap>" & rRs("Description") & "</Td>"
					CashProfitLossCol1 = "<TD class='no-top-right-border' Align=Left  nowrap>Cash PBT - Company</Td>"
					CashProfitLossCol2 = rRs("YTD")
					CashProfitLossCol3 = rRs("MBDG")
					CashProfitLossCol4 = rRs("PY1")
					CashProfitLossCol5 = rRs("PY2")
					CashProfitLossCol6 = rRs("PY3")
					CashProfitLossCol7 = rRs("PY4")
				End if 
				rRs.MoveNext
				
			Loop
			
			'
			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & SalesCol1
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(SalesCol2,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((SalesCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((SalesCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((SalesCol5),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((SalesCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right>"
			sSql = sSql & formatnumber((SalesCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & CommisionIncomeCol1
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((CommisionIncomeCol2),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((CommisionIncomeCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((CommisionIncomeCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((CommisionIncomeCol5),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((CommisionIncomeCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right>"
			sSql = sSql & formatnumber((CommisionIncomeCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & OtherIncome_LossCol1
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((OtherIncome_LossCol2),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((OtherIncome_LossCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((OtherIncome_LossCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((OtherIncome_LossCol5),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((OtherIncome_LossCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border' align=right>"
			sSql = sSql & formatnumber((OtherIncome_LossCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"


'response.write "<br>" &	SalesCol2 
'response.write "<br>" & CommisionIncomeCol2 
'response.write "<br>" & OtherIncome_LossCol2

			' "Total Revenue" printing start
			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' >"
			sSql = sSql & "Total Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((SalesCol4 + CommisionIncomeCol4 + OtherIncome_LossCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((SalesCol5 + CommisionIncomeCol5 + OtherIncome_LossCol5),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((SalesCol6 + CommisionIncomeCol6 + OtherIncome_LossCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border' align=right>"
			sSql = sSql & formatnumber((SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"


			' "Total Revenue" printing End




			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & PBTFromOperationCol1
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol2),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol5),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"



		' "PBT % on Total Revenue" printing start

			sSql = sSql & "<Tr Class=RepfontStyleItalicPrintReport>"
			sSql = sSql & "<TD class='no-top-right-border_Grey'>"
			sSql = sSql & "PBT % on Total Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol2/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol3/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol4/(SalesCol4 + CommisionIncomeCol4 + OtherIncome_LossCol4))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol5/(SalesCol5 + CommisionIncomeCol5 + OtherIncome_LossCol5))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol6 / (SalesCol6 + CommisionIncomeCol6 + OtherIncome_LossCol6))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTFromOperationCol7 / (SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "</Tr>"

			' "PBT % on Total Revenue" printing End


			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & InvestmentIncomeCol1
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((InvestmentIncomeCol2),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((InvestmentIncomeCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((InvestmentIncomeCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((InvestmentIncomeCol5),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((InvestmentIncomeCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right>"
			sSql = sSql & formatnumber((InvestmentIncomeCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & PBTCol1
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTCol2),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTCol5),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right>"
			sSql = sSql & formatnumber((PBTCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"


			' "Cash PBT - Operations" printing start

			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' bgcolor='#D8E4BC'>"
			sSql = sSql & "Cash PBT - Operations"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol2-InvestmentIncomeCol2,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol3-InvestmentIncomeCol3,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol4-InvestmentIncomeCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol5-InvestmentIncomeCol5,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol6-InvestmentIncomeCol6,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right bgcolor='#D8E4BC'>"
			sSql = sSql & formatnumber(CashProfitLossCol7-InvestmentIncomeCol7,0,,(-1)) 
			sSql = sSql & "</TD>"


			sSql = sSql & "</Tr>"

			' "Cash PBT - Operations" printing End


	' "Cash PBT % on Total Revenue" printing start

			sSql = sSql & "<Tr Class=RepfontStyleItalicPrintReport>"
			sSql = sSql & "<TD class='no-top-right-border_Grey'>"
			sSql = sSql & "Cash PBT % on Total Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol2-InvestmentIncomeCol2)/(SalesCol2 + CommisionIncomeCol2 + OtherIncome_LossCol2))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol3-InvestmentIncomeCol3)/(SalesCol3 + CommisionIncomeCol3 + OtherIncome_LossCol3))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol4-InvestmentIncomeCol4)/(SalesCol4 + CommisionIncomeCol4 + OtherIncome_LossCol4))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol5-InvestmentIncomeCol5)/(SalesCol5 + CommisionIncomeCol5 + OtherIncome_LossCol5))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol6-InvestmentIncomeCol6)/(SalesCol6 + CommisionIncomeCol6 + OtherIncome_LossCol6))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border_Grey' align=right>"
			sSql = sSql & formatnumber(((CashProfitLossCol7-InvestmentIncomeCol7)/(SalesCol7 + CommisionIncomeCol7 + OtherIncome_LossCol7))*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "</Tr>"

			' "Cash PBT % on Total Revenue" printing End


			sSql = sSql & "<Tr Class=RepContentFontPrintReport>"
			sSql = sSql & CashProfitLossCol1
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((CashProfitLossCol2),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((CashProfitLossCol3),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((CashProfitLossCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((CashProfitLossCol5),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber((CashProfitLossCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD class='no-top-border' align=right>"
			sSql = sSql & formatnumber((CashProfitLossCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			sSql = sSql & "</Table>"
			'		

			'		
			Response.write sSql 

		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (B)**********</Td>"
		End if		
		rRs.close
	End Function


	Private Function page2()


		sSql=" select "
'		sSql= sSql & " SUM(DECODE(DESCCD,'A5',ACT_MTD_BAL,0,'D5',ACT_MTD_BAL,0,'E6',ACT_MTD_BAL, 0)) FLASH , "
'		sSql= sSql & " SUM(DECODE(DESCCD,'A5',ACT_MTD_BAL_PY1,0,'D5',ACT_MTD_BAL_PY1,0,'E6',ACT_MTD_BAL_PY1, 0)) LY, "
'		sSql= sSql & " SUM(DECODE(DESCCD,'A5',ACT_MTD_BAL,0,'D5',ACT_MTD_BAL,0,'E6',ACT_MTD_BAL, 0)) - "
'		sSql= sSql & " SUM(DECODE(DESCCD,'A5',ACT_MTD_BAL_PY1,0,'D5',ACT_MTD_BAL_PY1,0,'E6',ACT_MTD_BAL_PY1, 0)) VS_LY, "
		sSql= sSql & " SUM(DECODE(DESCCD,'Q5',ACT_MTD_BAL,0)) FLASH , "
		sSql= sSql & " SUM(DECODE(DESCCD,'Q5',ACT_MTD_BAL_PY1,0)) LY, "
		sSql= sSql & " SUM(DECODE(DESCCD,'Q5',ACT_MTD_BAL,0)) - "
		sSql= sSql & " SUM(DECODE(DESCCD,'Q5',ACT_MTD_BAL_PY1,0)) VS_LY, "

		sSql= sSql & " B.ccompnamesn2, "
		'sSql= sSql & " SUM(DECODE(DESCCD,'A5',ACT_YTD_BAL,0,'D5',ACT_YTD_BAL,0,'E6',ACT_YTD_BAL, 0)) YearlyFLASH , "
		'sSql= sSql & " SUM(DECODE(DESCCD,'A5',ACT_YTD_BAL_PY1,0,'D5',ACT_YTD_BAL_PY1,0,'E6',ACT_YTD_BAL_PY1, 0)) YearlyLY, "
		'sSql= sSql & " SUM(DECODE(DESCCD,'A5',ACT_YTD_BAL,0,'D5',ACT_YTD_BAL,0,'E6',ACT_YTD_BAL, 0)) - "
		'sSql= sSql & " SUM(DECODE(DESCCD,'A5',ACT_YTD_BAL_PY1,0,'D5',ACT_YTD_BAL_PY1,0,'E6',ACT_YTD_BAL_PY1, 0)) YearlyVS_LY "
		sSql= sSql & " SUM(DECODE(DESCCD,'Q5',ACT_YTD_BAL, 0)) YearlyFLASH , "
		sSql= sSql & " SUM(DECODE(DESCCD,'Q5',ACT_YTD_BAL_PY1,0)) YearlyLY, "
		sSql= sSql & " SUM(DECODE(DESCCD,'Q5',ACT_YTD_BAL,0)) - "
		sSql= sSql & " SUM(DECODE(DESCCD,'Q5',ACT_YTD_BAL_PY1,0)) YearlyVS_LY "

		sSql= sSql & " from mistxn A , COMPANY B "
		sSql= sSql & " WHERE A.CCOMPCODE = B.CCOMPCODE "
		sSql= sSql & " AND cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
		sSql= sSql & " AND A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
		'sSql= sSql & " AND DESCCD IN ('A5', 'D5','E6') "
		sSql= sSql & " AND DESCCD IN ('A5', 'D5','E6','Q5') "
		sSql= sSql & " GROUP BY CYEAR, CMONTH,B.ccompnamesn2, CATEGORY "
		sSql= sSql & " ORDER BY 5 desc "

'response.write 	sSql



		rRs.Open Trim(sSql),cCn

		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""

		Totalcol1=0
		Totalcol2=0
		Totalcol3=0
		Totalcol4=0
		Totalcol5=0
		Totalcol6=0

		If Not rRs.Eof Then
			'
			sSql = ""

			sSql = sSql & "<Table border=0  width='100%' height='95%' cellPadding=1 cellspacing=0>"
			

			sSql = sSql & "<TR Class=RepfontStyle>"
			sSql = sSql & "<TD class='no-right-border' Align='center' rowspan=2>S.No.</td>"	
			sSql = sSql & "<TD class='no-right-border' Align='center' colspan=4 >Month " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
			sSql = sSql & "<TD class='no-right-border' Align='center' rowspan=2 width='30%'>Company</td>"	
			sSql = sSql & "<TD class='border-ALL' Align='center' colspan=4>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=RepfontStyle>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' colspan=2>vs LY</td>"	
			sSql = sSql & "</tr>"

			Do while not rRs.EOF
				'	
				SRLNO = SRLNO + 1
				Totalcol1 = Totalcol1 + rRs("FLASH")
				Totalcol2 = Totalcol2 + rRs("LY")
				Totalcol3 = Totalcol3 + rRs("VS_LY")

				Totalcol4 = Totalcol4 + rRs("YearlyFLASH")
				Totalcol5 = Totalcol5 + rRs("YearlyLY")
				Totalcol6 = Totalcol6 + rRs("YearlyVS_LY")

				sSql = sSql & "<TR Class=RepContentFont><TD class='no-top-right-border' Align='center' >" & SRLNO & "</td>"	

				If rRs("FLASH") <> 0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(rRs("FLASH"),0,,(-1))  & "</Td>"
				else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
				End if
				
				If rRs("LY") <> 0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-top-right-border' align='right'>" & formatnumber(rRs("LY"),0,,(-1)) & "</Td>"
				else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-top-right-border' align='right'>&nbsp;</Td>"
				End If 

				If rRs("FLASH") - rRs("LY") < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf rRs("FLASH") - rRs("LY") > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
				End if


				If rRs("VS_LY") <> 0  then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-right-left-top-border' align='right'>" & formatnumber(rRs("VS_LY"),0,,(-1)) & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-right-left-top-border' align='right'>&nbsp;</Td>"
				End if
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-top-right-border' align='left'><font class='RepContentFontNormal'>" & rRs("ccompnamesn2") & "</font></Td>"

				If rRs("YearlyFLASH") <> 0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-top-right-border' align='right'>" & formatnumber(rRs("YearlyFLASH"),0,,(-1)) & "</Td>"
				else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-top-right-border' align='right'>&nbsp;</Td>"
				End If
				
				If rRs("YearlyLY") <> 0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-top-right-border' align='right'>" & formatnumber(rRs("YearlyLY"),0,,(-1)) & "</Td>"
				else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  class='no-top-right-border' align='right'>&nbsp;</Td>"
				End If


				
				If rRs("YearlyFLASH") - rRs("YearlyLY") < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf rRs("YearlyFLASH") - rRs("YearlyLY") > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
				End if
				
				If rRs("YearlyVS_LY") <> 0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(rRs("YearlyVS_LY"),0,,(-1)) & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
				End If 

				sSql = sSql & "</TR>"	
				rRs.MoveNext
				
			Loop
			'
			sSql = sSql & "<TR Class=RepfontStyleCalculatedFiled>"
			sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
			If Totalcol1 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol1,0,,(-1)) & "</Td>"
			else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 
			'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol2,0,,(-1)) & "</Td>"
			If Totalcol2 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol2,0,,(-1)) & "</Td>"
			else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 


			If Totalcol1 - Totalcol2 < 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf Totalcol1 - Totalcol2 > 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
			End if
			
			'sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
			If Totalcol3 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(Totalcol3,0,,(-1)) & "</Td>"
			Else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='center'>Total</Td>"

			If Totalcol4 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol4,0,,(-1)) & "</Td>"
			else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If Totalcol5 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol5,0,,(-1)) & "</Td>"
			else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol4,0,,(-1)) & "</Td>"
			'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol5,0,,(-1)) & "</Td>"

			'sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
			If Totalcol4 - Totalcol5 < 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf Totalcol4 - Totalcol5 > 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
			End if

			If Totalcol6 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(Totalcol6,0,,(-1)) & "</Td>"
			else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
			End If 
			sSql = sSql & "</TR>"	

			sSql = sSql & "<TR Class=RepfontStyleReportRemark>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' colspan=10>*Revenue includes Sales, Commission and Other Income. <br>100% of the Financial Results of the above companies is included. "
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"	
			sSql = sSql & "</Table>"	
			Response.write sSql 

		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (C)**********</Td>"
		End if		
	rRs.close

	End function


	Private Function page3(PrintTableHeader)


'response.write sSql
'response.end



		rRs.Open Trim(sSql),cCn

		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""

		Totalcol1=0
		Totalcol2=0
		Totalcol3=0
		Totalcol4=0
		Totalcol5=0
		Totalcol6=0

		If Not rRs.Eof Then
			'
			sSql = ""

			'sSql = sSql & "<Table border=0  width='100%' height='95%' cellPadding=1 cellspacing=0>"
			sSql = sSql & "<Table border=0  width='100%'  cellPadding=1 cellspacing=0>"
			


'			sSql = sSql & "<TR Class=RepfontStyle>"
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' rowspan=2>Sr</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=4 >Month " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' rowspan=2 width='30%'>Company</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=4>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
'			sSql = sSql & "</tr>"

'			sSql = sSql & "<TR Class=RepfontStyle>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
'			sSql = sSql & "</tr>"

			If PrintTableHeader = "Y" then
				sSql = sSql & "<TR Class=RepfontStyle>"
				sSql = sSql & "<TD class='no-right-border' Align='center' rowspan=2>S.No.</td>"	
				sSql = sSql & "<TD class='no-right-border' Align='center' colspan=4 >Month " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
				sSql = sSql & "<TD class='no-right-border' Align='center' rowspan=2 width='30%'>Company</td>"	
				sSql = sSql & "<TD class='border-ALL' Align='center' colspan=4>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
				sSql = sSql & "</tr>"

				sSql = sSql & "<TR Class=RepfontStyle>"	
				sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
				sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
				sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
				sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
				sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
				sSql = sSql & "<TD class='no-top-border' Align='center' colspan=2>vs LY</td>"	
				sSql = sSql & "</tr>"

				sSql = sSql & "<TR Class=RepfontStyle>"	
				sSql = sSql & "<TD colspan=10 class='no-top-border' Align='center' >Section A</td>"	
				sSql = sSql & "</tr>"
			Else
				sSql = sSql & "<TR Class=RepfontStyle>"	
				sSql = sSql & "<TD colspan=10 class='no-top-border' Align='center' >Section B</td>"	
				sSql = sSql & "</tr>"
			
			End If 

			'sSql = sSql & "<TR Class=RepfontStyle><TD class='no-top-right-border' Align='center' >Sr</td>"	
			'sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			'sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year(LY)</td>"	
			'sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
			'sSql = sSql & "<TD class='no-top-right-border' Align='center' >Company</td>"	
			'sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			'sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year(LY)</td>"	
			'sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
			'sSql = sSql & "</tr>"

			Do while not rRs.EOF
				'	
				SRLNO = SRLNO + 1
				Totalcol1 = Totalcol1 + rRs("FLASH")
				Totalcol2 = Totalcol2 + rRs("LY")
				Totalcol3 = Totalcol3 + rRs("VS_LY")

				Totalcol4 = Totalcol4 + rRs("YearlyFLASH")
				Totalcol5 = Totalcol5 + rRs("YearlyLY")
				Totalcol6 = Totalcol6 + rRs("YearlyVS_LY")

				sSql = sSql & "<TR Class=RepContentFont><TD class='no-top-right-border' Align='center' width='6%' style='padding: 1px 2px 1px 2px;'>" & SRLNO & "</td>"	
				If rRs("FLASH") <0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='8%'><font color=#C00518>" & formatnumber(rRs("FLASH"),0,,(-1))  & "</font></Td>"
				ElseIf rRs("FLASH") >0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='8%'>" & formatnumber(rRs("FLASH"),0,,(-1))  & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='8%'>&nbsp;</Td>"
				End if

				If rRs("LY") <0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'><font color=#C00518>" & formatnumber(rRs("LY"),0,,(-1))  & "</font></Td>"
				ElseIf rRs("LY") >0 then 
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'>" & formatnumber(rRs("LY"),0,,(-1)) & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'>&nbsp;</Td>"
				End if
				
				
				

				If rRs("FLASH") - rRs("LY") < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='6%' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf rRs("FLASH") - rRs("LY") > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='6%' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center' width='6%'><img src='images/zero.png' width='60%'></Td>"
				End if


				If rRs("VS_LY") <> 0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-right-left-top-border' align='right' width='7%'>" & formatnumber(rRs("VS_LY"),0,,(-1)) & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-right-left-top-border' align='right' width='7%'>&nbsp;</Td>"
				End if
				
				sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='left'  width='30%'><font class='RepContentFontNormal'>" & rRs("ccompnamesn2") & "</font></Td>"

				If rRs("YearlyFLASH") <0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'><font color=#C00518>" & formatnumber(rRs("YearlyFLASH"),0,,(-1))  & "</font></Td>"
				ElseIf rRs("YearlyFLASH") >0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'>" & formatnumber(rRs("YearlyFLASH"),0,,(-1)) & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'>&nbsp;</Td>"
				End if

				If rRs("YearlyLY") <0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'><font color=#C00518>" & formatnumber(rRs("YearlyLY"),0,,(-1))  & "</font></Td>"
				ElseIf rRs("YearlyLY") >0 then 
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'>" & formatnumber(rRs("YearlyLY"),0,,(-1)) & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right' width='10%'>&nbsp;</Td>"
				End if


		

				If rRs("YearlyFLASH") - rRs("YearlyLY") < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0  width='6%' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf rRs("YearlyFLASH") - rRs("YearlyLY") > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0  width='6%' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center'  width='6%'><img src='images/zero.png' width='60%'></Td>"
				End if

				If rRs("YearlyVS_LY") <> 0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-left-border' align='right'>" & formatnumber(rRs("YearlyVS_LY"),0,,(-1)) & "</Td>"
				else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
				End if

				sSql = sSql & "</TR>"	
				rRs.MoveNext
				
			Loop
			'
			sSql = sSql & "<TR Class=RepfontStyleCalculatedFiled>"
			sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
			If Totalcol1 <> 0 Then
				If Totalcol1 < 0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'><font color=#C00518>" & formatnumber(Totalcol1,0,,(-1)) & "</font></Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol1,0,,(-1)) & "</Td>"
				End If 
			Else
				sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol2,0,,(-1)) & "</Td>"

			If Totalcol2 <> 0 then
				If Totalcol2 < 0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'><font color=#C00518>" & formatnumber(Totalcol2,0,,(-1)) & "</font></Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol2,0,,(-1)) & "</Td>"
				End If 
			Else
				sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 


				If Totalcol1 - Totalcol2 < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf Totalcol1 - Totalcol2 > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
				End if

			'sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
			If Totalcol3 <> 0 Then
				'If Totalcol3 < 0 then
				'	sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-right-left-top-border' align='right'><font color=#C00518>" & formatnumber(Totalcol3,0,,(-1)) & "</font></Td>"
				'Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(Totalcol3,0,,(-1)) & "</Td>"
				'End If 
			Else
				sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If
			If PrintTableHeader = "Y" Then
				sSql = sSql & "<TD style='padding: 1.5px 2px 1.5px 2px;' class='no-top-right-border' align='center'>Total - Section A</Td>"
			Else
				sSql = sSql & "<TD style='padding: 1.5px 2px 1.5px 2px;' class='no-top-right-border' align='center'>Total - Section B</Td>"
			End if
			
			If Totalcol4 <> 0 then
				If Totalcol4 < 0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'><font color=#C00518>" & formatnumber(Totalcol4,0,,(-1)) & "</font></Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol4,0,,(-1)) & "</Td>"
				End If 
			Else
				sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 
			'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol5,0,,(-1)) & "</Td>"

			If Totalcol5 <> 0 Then
				If Totalcol5 < 0 then
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'><font color=#C00518>" & formatnumber(Totalcol5,0,,(-1)) & "</font></Td>"
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol5,0,,(-1)) & "</Td>"
				End If 
			Else
				sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			'sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
			If Totalcol4 - Totalcol5 < 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf Totalcol4 - Totalcol5 > 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
			End if

			If Totalcol6 <> 0 then
				'If Totalcol6 < 0 then
				'	sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-left-border' align='right'><font color=#C00518>" & formatnumber(Totalcol6,0,,(-1)) & "</font></Td>"
				'Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-left-border' align='right'>" & formatnumber(Totalcol6,0,,(-1)) & "</Td>"
				'End If 

				
			Else
				sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
			End if
			sSql = sSql & "</TR>"

			If PrintTableHeader = "Y" Then
				SectionATotalCol1=Totalcol1
				SectionATotalCol2=Totalcol2
				SectionATotalCol3=Totalcol3
				SectionATotalCol4=Totalcol4
				SectionATotalCol5=Totalcol5
				SectionATotalCol6=Totalcol6
			End If
			

			If PrintTableHeader = "N" Then
				sSql = sSql & "<TR Class=RepfontStyleCalculatedFiled>"
				sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
				If (SectionATotalCol1+Totalcol1) <> 0 then
					If (SectionATotalCol1+Totalcol1) < 0 then
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'><font color=#C00518>" & formatnumber(SectionATotalCol1+Totalcol1,0,,(-1)) & "</font></Td>"
					Else
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>" & formatnumber(SectionATotalCol1+Totalcol1,0,,(-1)) & "</Td>"
					End If 
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
				End If 

				'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol2,0,,(-1)) & "</Td>"

				If (SectionATotalCol2 + Totalcol2) <> 0 then
					If (SectionATotalCol2+Totalcol2) < 0 then
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'><font color=#C00518>" & formatnumber(SectionATotalCol2 + Totalcol2,0,,(-1)) & "</font></Td>"
					Else
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>" & formatnumber(SectionATotalCol2 + Totalcol2,0,,(-1)) & "</Td>"
					End If 
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
				End If 


				If ((SectionATotalcol1 - SectionATotalcol2) + (Totalcol1 - Totalcol2)) < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf ((SectionATotalcol1 - SectionATotalcol2) + (Totalcol1 - Totalcol2)) > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
				End if

				'sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
				If (SectionATotalcol3 + Totalcol3) <> 0 then
					'If (SectionATotalCol3+Totalcol3) < 0 then
					'	sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-right-left-top-border' align='right'><font color=#C00518>" & formatnumber((SectionATotalcol3 + Totalcol3),0,,(-1)) & "</font></Td>"
					'Else
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-right-left-top-border' align='right'>" & formatnumber((SectionATotalcol3 + Totalcol3),0,,(-1)) & "</Td>"
					'End If 
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
				End If
				
				sSql = sSql & "<TD style='padding: 1.5px 2px 1.5px 2px;' class='no-top-right-border' align='center'>Sum Total Section A + B</Td>"
				
				If (SectionATotalcol4 + Totalcol4) <> 0 then
					If (SectionATotalCol4+Totalcol4) < 0 then
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'><font color=#C00518>" & formatnumber((SectionATotalcol4 + Totalcol4),0,,(-1)) & "</font></Td>"
					Else
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>" & formatnumber((SectionATotalcol4 + Totalcol4),0,,(-1)) & "</Td>"
					End If 
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
				End If 
				'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol5,0,,(-1)) & "</Td>"

				If (SectionATotalcol5 + Totalcol5) <> 0 then
					If (SectionATotalCol5+Totalcol5) < 0 then
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'><font color=#C00518>" & formatnumber((SectionATotalcol5 + Totalcol5),0,,(-1)) & "</font></Td>"
					Else
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>" & formatnumber((SectionATotalcol5 + Totalcol5),0,,(-1)) & "</Td>"
					End If 
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
				End If 

				'sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
				If ((SectionATotalcol4 - SectionATotalcol5) + (Totalcol4 - Totalcol5)) < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf ((SectionATotalcol4 - SectionATotalcol5) + (Totalcol4 - Totalcol5)) > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
				End if

				If (SectionATotalcol6 + Totalcol6) <> 0 then
					'If (SectionATotalCol6+Totalcol6) < 0 then
					'	sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-left-border' align='right'><font color=#C00518>" & formatnumber((SectionATotalcol6 + Totalcol6),0,,(-1)) & "</font></Td>"
					'Else
						sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-left-border' align='right'>" & formatnumber((SectionATotalcol6 + Totalcol6),0,,(-1)) & "</Td>"
					'End If 
				Else
					sSql = sSql & "<TD style='padding: 1px 2px 1px 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
				End if
				sSql = sSql & "</TR>"
				
				
				'
				sSql = sSql & "<TR Class=RepfontStyleReportRemark>"
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' colspan=10>100% of the Financial Results of the above companies is included. "
				sSql = sSql & "</TD>"
				sSql = sSql & "</TR>"	
			End if

			sSql = sSql & "</Table>"
			Response.write sSql 

		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (D)**********</Td>"
		End if		
		rRs.close

	End function


Private Function page4()


		sSql=" select "
		sSql= sSql & " SUM(DECODE(DESCCD,'M5',ACT_MTD_BAL,0) - DECODE(DESCCD,'E5',ACT_MTD_BAL,0)) FLASH , "
		sSql= sSql & " SUM(DECODE(DESCCD,'M5',ACT_MTD_BAL_PY1,0) - DECODE(DESCCD,'E5',ACT_MTD_BAL_PY1,0)) LY, "
		sSql= sSql & " SUM(DECODE(DESCCD,'M5',ACT_MTD_BAL,0) - DECODE(DESCCD,'E5',ACT_MTD_BAL,0)) - "
		sSql= sSql & " SUM(DECODE(DESCCD,'M5',ACT_MTD_BAL_PY1,0) - DECODE(DESCCD,'E5',ACT_MTD_BAL_PY1,0)) VS_LY, "
		sSql= sSql & " B.ccompnamesn2, "
		sSql= sSql & " SUM(DECODE(DESCCD,'M5',ACT_YTD_BAL,0) - DECODE(DESCCD,'E5',ACT_YTD_BAL,0) ) YearlyFLASH , "
		sSql= sSql & " SUM(DECODE(DESCCD,'M5',ACT_YTD_BAL_PY1,0) - DECODE(DESCCD,'E5',ACT_YTD_BAL_PY1,0)) YearlyLY, "
		sSql= sSql & " SUM(DECODE(DESCCD,'M5',ACT_YTD_BAL,0) - DECODE(DESCCD,'E5',ACT_YTD_BAL,0)) - "
		sSql= sSql & " SUM(DECODE(DESCCD,'M5',ACT_YTD_BAL_PY1,0) - DECODE(DESCCD,'E5',ACT_YTD_BAL_PY1,0)) YearlyVS_LY "
		sSql= sSql & " from mistxn A , COMPANY B "
		sSql= sSql & " WHERE A.CCOMPCODE = B.CCOMPCODE "
		sSql= sSql & " AND cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
		sSql= sSql & " AND A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
		sSql= sSql & " AND DESCCD IN ('E5','M5') "
		sSql= sSql & " GROUP BY CYEAR, CMONTH,B.ccompnamesn2, CATEGORY "
		sSql= sSql & " ORDER BY 5 desc "

'		sSql=" select "
'		sSql= sSql & " SUM(ACT_MTD_BAL) FLASH , "
'		sSql= sSql & " SUM(ACT_MTD_BAL_PY1) LY, "
'		sSql= sSql & " SUM(ACT_MTD_BAL) - "
'		sSql= sSql & " SUM(ACT_MTD_BAL_PY1) VS_LY, "
'		sSql= sSql & " A.CCOMPCODE, "
'		sSql= sSql & " SUM(ACT_YTD_BAL) YearlyFLASH , "
'		sSql= sSql & " SUM(ACT_YTD_BAL_PY1) YearlyLY, "
'		sSql= sSql & " SUM(ACT_YTD_BAL) - "
'		sSql= sSql & " SUM(ACT_YTD_BAL_PY1) YearlyVS_LY "
'		sSql= sSql & " from mistxn A , COMPANY B "
'		sSql= sSql & " WHERE A.CCOMPCODE = B.CCOMPCODE "
'		sSql= sSql & " AND cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
'		sSql= sSql & " AND A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
'		sSql= sSql & " AND DESCCD IN ('M5') "
'		sSql= sSql & " GROUP BY CYEAR, CMONTH,A.CCOMPCODE, CATEGORY "
'		sSql= sSql & " ORDER BY 5 desc "

	



		rRs.Open Trim(sSql),cCn

		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""

		Totalcol1=0
		Totalcol2=0
		Totalcol3=0
		Totalcol4=0
		Totalcol5=0
		Totalcol6=0

		If Not rRs.Eof Then
			'
			sSql = ""

			sSql = sSql & "<Table border=0  width='100%' height='95%'cellPadding=1 cellspacing=0>"
		

'			sSql = sSql & "<TR Class=RepfontStyle>"
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' rowspan=2>Sr</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=4 >Month " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' rowspan=2 width='30%'>Company</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=4>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
'			sSql = sSql & "</tr>"
'
'			sSql = sSql & "<TR Class=RepfontStyle>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
'			sSql = sSql & "</tr>"
'



			sSql = sSql & "<TR Class=RepfontStyle>"
			sSql = sSql & "<TD class='no-right-border' Align='center' rowspan=2>S.No.</td>"	
			sSql = sSql & "<TD class='no-right-border' Align='center' colspan=4 >Month " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
			sSql = sSql & "<TD class='no-right-border' Align='center' rowspan=2 >Company</td>"	
			sSql = sSql & "<TD class='border-ALL' Align='center' colspan=4>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=RepfontStyle>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year<br>(LY)</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' colspan=2>vs LY</td>"	
			sSql = sSql & "</tr>"


'			sSql = sSql & "<TR Class=RepfontStyle><TD class='no-top-right-border' Align='center' >Sr</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year(LY)</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Company</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' >Last Year(LY)</td>"	
'			sSql = sSql & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
'			sSql = sSql & "</tr>"

			Do while not rRs.EOF
				'	
				SRLNO = SRLNO + 1
				Totalcol1 = Totalcol1 + rRs("FLASH")
				Totalcol2 = Totalcol2 + rRs("LY")
				Totalcol3 = Totalcol3 + rRs("VS_LY")

				Totalcol4 = Totalcol4 + rRs("YearlyFLASH")
				Totalcol5 = Totalcol5 + rRs("YearlyLY")
				Totalcol6 = Totalcol6 + rRs("YearlyVS_LY")

				sSql = sSql & "<TR Class=RepContentFont><TD class='no-top-right-border' Align='center'  width='6%'>" & SRLNO & "</td>"	

				If rRs("FLASH") <0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='8%'><font color=#C00518>" & formatnumber(rRs("FLASH"),0,,(-1))  & "</font></Td>"
				ElseIf rRs("FLASH") >0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='8%'>" & formatnumber(rRs("FLASH"),0,,(-1))  & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='8%'>&nbsp;</Td>"
				End if


				If rRs("LY") <0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'><font color=#C00518>" & formatnumber(rRs("LY"),0,,(-1))  & "</font></Td>"
				ElseIf rRs("LY") >0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'>" & formatnumber(rRs("LY"),0,,(-1))  & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'>&nbsp;</Td>"
				End if

				'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(rRs("LY"),0,,(-1)) & "</Td>"

				If rRs("FLASH") - rRs("LY") < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0  width='6%' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf rRs("FLASH") - rRs("LY") > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0  width='6%' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center'  width='6%'><img src='images/zero.png' width='60%'></Td>"
				End if


				'If rRs("VS_LY") <0 then
				'	sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'><font color=#C00518>" & formatnumber(rRs("VS_LY"),0,,(-1))  & "</font></Td>"
				'Else 
				'	sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(rRs("VS_LY"),0,,(-1))  & "</Td>"
				'End if
				If rRs("VS_LY") <> 0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right' width='7%'>" & formatnumber(rRs("VS_LY"),0,,(-1)) & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right' width='7%'>&nbsp;</Td>"
				End If 
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='left' width='30%'><font class='RepContentFontNormal'>" & rRs("ccompnamesn2") & "</font></Td>"


				If rRs("YearlyFLASH") <0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'><font color=#C00518>" & formatnumber(rRs("YearlyFLASH"),0,,(-1))  & "</font></Td>"
				ElseIf rRs("YearlyFLASH") >0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'>" & formatnumber(rRs("YearlyFLASH"),0,,(-1))  & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'>&nbsp;</Td>"
				End if
				'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(rRs("YearlyFLASH"),0,,(-1)) & "</Td>"


				If rRs("YearlyLY") <0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'><font color=#C00518>" & formatnumber(rRs("YearlyLY"),0,,(-1))  & "</font></Td>"
				ElseIf rRs("YearlyLY") > 0 then 
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'>" & formatnumber(rRs("YearlyLY"),0,,(-1))  & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' width='10%'>&nbsp;</Td>"
				End If
				
				'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(rRs("YearlyLY"),0,,(-1)) & "</Td>"

				If rRs("YearlyFLASH") - rRs("YearlyLY") < 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0  width='6%' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
				ElseIf rRs("YearlyFLASH") - rRs("YearlyLY") > 0 Then
					sSql = sSql & "<TD class='no-top-right-border' Align=center border=0  width='6%' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
				Else
					sSql = sSql & "<TD class='no-top-right-border' align='center' width='6%'><img src='images/zero.png' width='60%'></Td>"
				End if

				'If rRs("YearlyVS_LY") <0 then
				'	sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'><font color=#C00518>" & formatnumber(rRs("YearlyVS_LY"),0,,(-1))  & "</font></Td>"
				'Else 
				'	sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(rRs("YearlyVS_LY"),0,,(-1))  & "</Td>"
				'End if

				If rRs("YearlyVS_LY") <> 0 then
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(rRs("YearlyVS_LY"),0,,(-1)) & "</Td>"
				Else
					sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
				End If 

				sSql = sSql & "</TR>"	
				rRs.MoveNext
				
			Loop
			'
			sSql = sSql & "<TR Class=RepfontStyleCalculatedFiled>"
			sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
			If Totalcol1 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol1,0,,(-1)) & "</Td>"
			Else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol2,0,,(-1)) & "</Td>"

			If Totalcol2 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol2,0,,(-1)) & "</Td>"
			Else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If Totalcol1 - Totalcol2 < 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf Totalcol1 - Totalcol2 > 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
			End if

			'sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"
			If Totalcol3 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(Totalcol3,0,,(-1)) & "</Td>"
			Else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 

			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='center'>Total</Td>"

			If Totalcol4 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol4,0,,(-1)) & "</Td>"
			Else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 
		
			If Totalcol5 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol5,0,,(-1)) & "</Td>"
			Else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 
			'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol4,0,,(-1)) & "</Td>"
			'sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(Totalcol5,0,,(-1)) & "</Td>"
			'sSql = sSql & "<TD class='no-top-right-border' align='right'>&nbsp;</Td>"

			If Totalcol4 - Totalcol5 < 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf Totalcol4 - Totalcol5 > 0 Then
				sSql = sSql & "<TD class='no-top-right-border' Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSql = sSql & "<TD class='no-top-right-border' align='center'><img src='images/zero.png' width='60%'></Td>"
			End if

			If Totalcol6 <> 0 then
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(Totalcol6,0,,(-1)) & "</Td>"
			Else
				sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
			End If 
			sSql = sSql & "</TR>"	
			sSql = sSql & "<TR Class=RepfontStyleReportRemark>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;'  colspan=10>100% of the Financial Results of the above companies is included. "
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"	

			sSql = sSql & "</Table>"
			Response.write sSql 

		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (E)**********</Td>"
		End if		
		rRs.close
	End function



	Private Function page7_Top10_Bottom10()

		sSqlQry=  " select * from "
		sSqlQry= sSqlQry & " (select "
		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL,0)) FLASH , "
'		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL_PY1,0)) LY, "
'		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL,0)) - SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL_PY1,0)) VS_LY, "
		sSqlQry= sSqlQry & " B.ccompnamesn2, "
		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL,0)) YearlyFLASH  "
'		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL_PY1,0)) YearlyLY, "
'		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL,0)) - SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL_PY1,0)) YearlyVS_LY "
		sSqlQry= sSqlQry & " from mistxn A , COMPANY B "
		sSqlQry= sSqlQry & " WHERE A.CCOMPCODE = B.CCOMPCODE "
		sSqlQry= sSqlQry & " AND cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
		sSqlQry= sSqlQry & " AND A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
		sSqlQry= sSqlQry & " AND DESCCD IN ('R5') "
		sSqlQry= sSqlQry & " GROUP BY CYEAR, CMONTH,B.ccompnamesn2, CATEGORY "
		sSqlQry= sSqlQry & " ORDER BY 3 desc ) x"
		sSqlQry= sSqlQry & " where rownum<=10 "




		rRs.Open Trim(sSqlQry),cCn

		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""

		Totalcol1=0
		sSql = ""
		sSqlT1= ""
		sSqlT2= ""
		sSqlT3= ""

		If Not rRs.Eof Then
			'


			sSqlT1 = sSqlT1 & "<Table border=0  width='50%' cellPadding=1 cellspacing=0 >"
			
			sSqlT1 = sSqlT1 & "<TR Class=GreenTop10Page7><TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' Align='center' colspan=3 nowrap>Top 10 Profit Companies (PBT-Operations)</td></TR>"	

			sSqlT1 = sSqlT1 & "<TR Class=RepfontStylePage7><TD class='no-top-right-border' Align='center' nowrap >S.No.</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >Company</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-border' Align='center' nowrap >&nbsp;Flash YTD&nbsp;<br>" &  MonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "&nbsp;</td>"	
			sSqlT1 = sSqlT1 & "</tr>"

			Do while not rRs.EOF
				'	
				SRLNO = SRLNO + 1

				Totalcol1 = Totalcol1 + rRs("YearlyFLASH")

				sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalPage7><TD class='no-top-right-border_Grey' Align='center' >" & SRLNO & "</td>"	
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='left' nowrap>" & rRs("ccompnamesn2") & "&nbsp;</Td>"
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-border_Grey' align='right'>" & formatnumber(rRs("YearlyFLASH"),0,,(-1)) & "</Td>"

				sSqlT1 = sSqlT1 & "</TR>"	
				rRs.MoveNext
				
			Loop
			
			'
			sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleCalculatedFiledPrintReportPage7>"
			sSqlT1 = sSqlT1 & "<TD class='no-right-border' align='right'>&nbsp;</Td>"
			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-border' align='left'>Total</Td>"
			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='border-ALL' align='right'>" & formatnumber(Totalcol1,0,,(-1)) & "</Td>"
			sSqlT1 = sSqlT1 & "</TR></table>"	
		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (F)**********</Td>"
		End if		

		rRs.close

		sSqlQry=  " select * from "
		sSqlQry= sSqlQry & " (select "
		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL,0)) FLASH , "
'		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL_PY1,0)) LY, "
'		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL,0)) - SUM(DECODE(DESCCD,'R5',ACT_MTD_BAL_PY1,0)) VS_LY, "
		sSqlQry= sSqlQry & " B.ccompnamesn2, "
		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL,0)) YearlyFLASH  "
'		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL_PY1,0)) YearlyLY, "
'		sSqlQry= sSqlQry & " SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL,0)) - SUM(DECODE(DESCCD,'R5',ACT_YTD_BAL_PY1,0)) YearlyVS_LY "
		sSqlQry= sSqlQry & " from mistxn A , COMPANY B "
		sSqlQry= sSqlQry & " WHERE A.CCOMPCODE = B.CCOMPCODE "
		sSqlQry= sSqlQry & " AND cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' "
		sSqlQry= sSqlQry & " AND A.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
		sSqlQry= sSqlQry & " AND DESCCD IN ('R5') "
		sSqlQry= sSqlQry & " GROUP BY CYEAR, CMONTH,B.ccompnamesn2, CATEGORY "
		sSqlQry= sSqlQry & " ORDER BY 3 asc ) x"
		sSqlQry= sSqlQry & " where rownum<=10 "





		rRs.Open Trim(sSqlQry),cCn

		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""

		Totalcol1=0

		If Not rRs.Eof Then
			'
			sSqlT2 = sSqlT2 & "<Table border=0  width='50%' cellPadding=1 cellspacing=0>"
			
			sSqlT2 = sSqlT2 & "<TR Class=RedTransparentHeadingSmallPage7><TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' Align='center' colspan=3 nowrap>Top 10 Loss Companies (PBT-Operations)</td></TR>"	

			sSqlT2 = sSqlT2 & "<TR Class=RepfontStylePage7><TD class='no-top-right-border' Align='center' nowrap>S.No.</td>"	
			sSqlT2 = sSqlT2 & "<TD class='no-top-right-border' Align='center' >Company</td>"	
			sSqlT2 = sSqlT2 & "<TD class='no-top-border' Align='center' nowrap >&nbsp;Flash YTD&nbsp;<br>" & MonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
			sSqlT2 = sSqlT2 & "</tr>"

			Do while not rRs.EOF
				'	
				SRLNO = SRLNO + 1

				Totalcol1 = Totalcol1 + rRs("YearlyFLASH")

				sSqlT2 = sSqlT2 & "<TR Class=RepContentFontNormalPage7><TD class='no-top-right-border_Grey' Align='Center' >" & SRLNO & "</td>"	
				sSqlT2 = sSqlT2 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='left' nowrap>" & rRs("ccompnamesn2") & "&nbsp;</Td>"
				sSqlT2 = sSqlT2 & "<TD style='padding: 0 2px 0 2px;' class='no-top-border_Grey' align='right'>" & formatnumber(rRs("YearlyFLASH"),0,,(-1)) & "</Td>"

				sSqlT2 = sSqlT2 & "</TR>"	
				rRs.MoveNext
				
			Loop
			
			'
			sSqlT2 = sSqlT2 & "<TR Class=RepfontStyleCalculatedFiledPrintReportPage7>"
			sSqlT2 = sSqlT2 & "<TD class='no-right-border' align='right'>&nbsp;</Td>"
			sSqlT2 = sSqlT2 & "<TD style='padding: 0 2px 0 2px;' class='no-right-border' align='left'>Total</Td>"
			sSqlT2 = sSqlT2 & "<TD style='padding: 0 2px 0 2px;' class='border-ALL' align='right'>" & formatnumber(Totalcol1,0,,(-1)) & "</Td>"
			sSqlT2 = sSqlT2 & "</TR></table>"	
		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (G)**********</Td>"
		End if		
		rRs.close


		sSqlQry=  " select c.ccompnamesn2 ,sum(a.ACT_MTD_BAL) col2 ,"
		sSqlQry= sSqlQry & " Min((select act_mtd_bal from mistxn where ccompcode= a.ccompcode and desccd='E5' and cyear='" & TheReportYear & "' and cmonth='" & TheReportMonth & "')) col1,  "
		sSqlQry= sSqlQry & " sum(a.ACT_YTD_BAL) col4 ,"
		sSqlQry= sSqlQry & " Min((select act_ytd_bal from mistxn where ccompcode= a.ccompcode and desccd='E5' and cyear='" & TheReportYear & "' and cmonth='" & TheReportMonth & "')) col3  "
		sSqlQry= sSqlQry & " from DIVIDEND_INCOME_NEW a, company c"
		sSqlQry= sSqlQry & " where a.ccompcode=c.ccompcode and a.cyear='" & TheReportYear & "' and a.cmonth = '" & TheReportMonth & "' "
		sSqlQry= sSqlQry & " group by c.ccompnamesn2  "

		sSqlQry= sSqlQry & " Union "
		sSqlQry= sSqlQry & " select c.ccompnamesn2,0,a.act_mtd_bal,0,a.act_ytd_bal from mistxn  a, company c where  a.ccompcode=c.ccompcode and a.desccd='E5' and a.cyear='" & TheReportYear & "' and a.cmonth='" & TheReportMonth & "' "
		sSqlQry= sSqlQry & " and (a.act_mtd_bal<>0 or a.act_ytd_bal<>0) and (a.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') and  a.ccompcode not in (select distinct ccompcode from DIVIDEND_INCOME_NEW where cyear='" & TheReportYear & "' and cmonth='" & TheReportMonth & "'))"
		sSqlQry= sSqlQry & " order by 1"


'response.write sSqlQry
'response.end
		rRs.Open Trim(sSqlQry),cCn

		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""

		Totalcol1=0
		Totalcol2=0
		Totalcol3=0
		Totalcol4=0

		If Not rRs.Eof Then
			'

			sSqlT3 = sSqlT3 & "<br>"
			sSqlT3 = sSqlT3 & "<Table border=0  width='100%'  cellPadding=1 cellspacing=0 >"
			
			sSqlT3 = sSqlT3 & "<TR Class=GreenTransparentHeadingNormalPage7><TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' Align='left' colspan=8 nowrap>Investment Income & Inter-Company Dividends</td></TR>"	

			sSqlT3 = sSqlT3 & "<TR Class=RepfontStyle><TD class='no-top-right-border' Align='center' >&nbsp;</td>"	
			sSqlT3 = sSqlT3 & "<TD class='no-top-right-border' Align='center' colspan=3 >MTD " & MonthsArr(cint(TheReportMonth)-1) & "&nbsp" & TheReportYear & "</td>"	
			sSqlT3 = sSqlT3 & "<TD class='no-top-border_thick' Align='center' colspan=3 >YTD " & MonthsArr(cint(TheReportMonth)-1) & "&nbsp" & TheReportYear & "</td>"	
			sSqlT3 = sSqlT3 & "</tr>"
			sSqlT3 = sSqlT3 & "<TR Class=RepfontStylePage7>"
			'sSqlT3 = sSqlT3 & "<TD class='no-top-right-border' Align='center'>Sr</td>"	
			sSqlT3 = sSqlT3 & "<TD class='no-top-right-border' Align='center' >Company</td>"	
			sSqlT3 = sSqlT3 & "<TD class='no-top-right-border' Align='center' >Investment <br> Income (A) </td>"
			sSqlT3 = sSqlT3 & "<TD class='no-top-right-border' Align='center' >Inter-Co.<br> Dividend(B) </td>"	
			sSqlT3 = sSqlT3 & "<TD class='no-top-right-border' Align='center' nowrap>&nbsp;Net<br>&nbsp;(A-B)&nbsp;&nbsp;</td>"	
			sSqlT3 = sSqlT3 & "<TD class='no-top-right-border_thick' Align='center' >Investment <br> Income (A) </td>"
			sSqlT3 = sSqlT3 & "<TD class='no-top-right-border' Align='center' >Inter-Co.<br> Dividend(B) </td>"	
			sSqlT3 = sSqlT3 & "<TD class='no-top-border' Align='center' >&nbsp;Net<br>&nbsp;(A-B)&nbsp;&nbsp;</td>"	
			sSqlT3 = sSqlT3 & "</tr>"

			Do while not rRs.EOF
				'	
				If rRs("col1")<>0 or rRs("col2")<>0 or rRs("col3")<>0 or rRs("col4")<>0 then
					SRLNO = SRLNO + 1

					Totalcol1 = Totalcol1 + rRs("col1")
					Totalcol2 = Totalcol2 + rRs("col2")
					Totalcol3 = Totalcol3 + rRs("col3")
					Totalcol4 = Totalcol4 + rRs("col4")
					TotalGrp1=TotalGrp1 + (rRs("col1")-rRs("col2"))
					TotalGrp2=TotalGrp2 + (rRs("col3")-rRs("col4"))

					sSqlT3 = sSqlT3 & "<TR Class=RepContentFontNormalPage7 height='20'>"
					'sSqlT3 = sSqlT3 & "<TD class='no-top-right-border_Grey' Align='center' >" & SRLNO & "</td>"	
					sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='left' nowrap>" & rRs("ccompnamesn2") & "&nbsp;</Td>"
					If rRs("col1") <> 0 then
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='right'>" & formatnumber(rRs("col1"),0,,(-1)) & "</Td>"
					Else
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='right'>&nbsp;</Td>"
					End if 

					If rRs("col2") <> 0 then
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='right'>" & formatnumber(rRs("col2"),0,,(-1)) & "</Td>"
					Else
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='right'>&nbsp;</Td>"
					End if 

					If rRs("col1")-rRs("col2") <> 0 then
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='right'>" &  formatnumber(rRs("col1")-rRs("col2"),0,,(-1)) & "</Td>"
					Else
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='right'>&nbsp;</Td>"
					End if 

					If rRs("col3") <> 0 then
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey_thick' align='right'>" & formatnumber(rRs("col3"),0,,(-1)) & "</Td>"
					Else
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey_thick' align='right'>&nbsp</Td>"
					End if 


					If rRs("col4") <> 0 then
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='right'>" & formatnumber(rRs("col4"),0,,(-1)) & "</Td>"
					Else
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_Grey' align='right'>&nbsp;</Td>"
					End if 
					

					If rRs("col3")-rRs("col4") <> 0 then
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-border_Grey' align='right'>" & formatnumber(rRs("col3")-rRs("col4"),0,,(-1)) & "</Td>"
					Else
						sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-top-border_Grey' align='right'>&nbsp;</Td>"
					End if 
					
					
					sSqlT3 = sSqlT3 & "</TR>"
				End If 
				rRs.MoveNext
				
			Loop
		
			'
			sSqlT3 = sSqlT3 & "<TR Class='RepfontStyleCalculatedFiledPrintReportPage7' height='20'>"
			'sSqlT3 = sSqlT3 & "<TD class='no-right-border' align='right'>&nbsp;</Td>"
			sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-right-border' align='left'>Total</Td>"
			sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-right-border' align='right'>" & formatnumber(Totalcol1,0,,(-1)) & "</Td>"
			sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-right-border' align='right'>" & formatnumber(Totalcol2,0,,(-1)) & "</Td>"
			sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-right-border' align='right'>" & formatnumber(TotalGrp1,0,,(-1)) & "</Td>"
			sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-right-border_Grey_thick' align='right'>" & formatnumber(Totalcol3,0,,(-1)) & "</Td>"
			sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='no-right-border' align='right'>" & formatnumber(Totalcol4,0,,(-1)) & "</Td>"
			sSqlT3 = sSqlT3 & "<TD style='padding: 0 2px 0 2px;' class='border-ALL' align='right'>" & formatnumber(TotalGrp2,0,,(-1)) & "</Td>"
			sSqlT3 = sSqlT3 & "</TR></table>"	
		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (H)**********</Td>"
		End if	
		rRs.close





			sSql = sSql & "<Table border=0  width='100%' cellPadding=0 cellspacing=0 >"

'			sSql = sSql & "</TR>"
'			sSql = sSql & "<TR>"
'			sSql = sSql & "<TD class='no-top-right-border' align=left colspan=2>&nbsp;"
'			sSql = sSql & "</TD>"
'			sSql = sSql & "<TR>"

			sSql = sSql & "<TR>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' align=left width='50%'>"
			sSql = sSql & sSqlT1 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' align=right>"
			sSql = sSql & sSqlT2
			sSql = sSql & "</TD>"

			sSql = sSql & "</TR>"
			sSql = sSql & "<TR>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' align=left colspan=2>&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TR>"

			sSql = sSql & "</TR>"
			sSql = sSql & "<TR>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' align=left colspan=2>&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TR>"

			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' align=left colspan=2>"
			sSql = sSql & sSqlT3 
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"


			sSql = sSql & "</Table>"
			Response.write sSql 


	End function

	Private Function page8_9_10_11_12_13(sSql1,sSql2,Title1,Title2,period,textSize)




		rRs.Open Trim(sSql1),cCn

		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""

		Totalcol1=0
		Totalcol2=0
		Totalcol3=0
		Totalcol4=0
		Totalcol5=0
		Totalcol6=0
		Totalcol7=0
		Totalcol8=0
		Totalcol9=0
		Totalcol10=0
		Totalcol11=0
		Totalcol12=0


		sSql = ""
		sSqlT1= ""

		If Not rRs.Eof Then
			'

			If textSize="SMALL" Then
				sSqlT1 = sSqlT1 & "<Table border=0  width='100%' cellPadding=0 cellspacing=0>"
			Else
				sSqlT1 = sSqlT1 & "<Table border=0  width='100%' cellPadding=1 cellspacing=0>"
			End If

'			sSqlT1 = sSqlT1 & "<Table border=1  width='100%' cellPadding=1 cellspacing=0>"


'			If textSize="SMALL" then
'				sSqlT1 = sSqlT1 & "<TR Class=AllReportHeaderSMALL><TD class='no-top-right-border' Align='center' colspan=14 nowrap>" & Title1 & "</td><TD class='no-top-right-border' colspan=4 align=right>Amount (RO '000)</td></TR>"	
'			Else
'				sSqlT1 = sSqlT1 & "<TR Class=AllReportHeader><TD class='no-top-right-border' Align='center' colspan=14 nowrap>" & Title1 & "</td><TD class='no-top-right-border' colspan=4 align=right>Amount (RO '000)</td></TR>"	
'			End If 

			If period = "M" Then
				If textSize="SMALL" then
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
				Else
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStylePage8>"	
				End if
				sSqlT1 = sSqlT1 & "<TD class='no-right-border' Align='center' rowspan=2 >Company</td>"	
				sSqlT1 = sSqlT1 & "<TD class='no-right-border' Align='center' colspan=4>MTD Flash " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
				sSqlT1 = sSqlT1 & "<TD class='no-right-border_thick' Align='center' colspan=4>MTD Budget " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
				sSqlT1 = sSqlT1 & "<TD class='border-ALL_left_thick' Align='center' colspan=8>MTD Flash vs Budget</td>"	
				sSqlT1 = sSqlT1 & "</tr>"
			ElseIf period="Y" Then
				If textSize="SMALL" then
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
				Else
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStylePage8>"	
				End if

				sSqlT1 = sSqlT1 & "<TD class='no-right-border' Align='center' rowspan=2>Company</td>"	
				sSqlT1 = sSqlT1 & "<TD class='no-right-border' Align='center' colspan=4>YTD Flash " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
				sSqlT1 = sSqlT1 & "<TD class='no-right-border_thick' Align='center' colspan=4>YTD Budget " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
				sSqlT1 = sSqlT1 & "<TD class='border-ALL_left_thick' Align='center' colspan=8>YTD Flash vs Budget</td>"	
				sSqlT1 = sSqlT1 & "</tr>"
			End If 

			If textSize="SMALL" then
				sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
			Else
				sSqlT1 = sSqlT1 & "<TR Class=RepfontStylePage8>"	
			End if

			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >Revenue</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' nowrap >PBT Ops</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >Invst<br>Income</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >PBT<br>Company</td>"	

			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border_thick' Align='center' >Revenue</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' nowrap >PBT Ops</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >Invst<br>Income</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >PBT<br>Company</td>"	
			
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center'></Td>"
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border_thick' Align='center' colspan=2 >Revenue</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' ></Td>"
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' nowrap colspan=2>PBT Ops</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' ></Td>"
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' colspan=2>Invst<br>Income</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' ></Td>"
			sSqlT1 = sSqlT1 & "<TD class='no-top-border' Align='center' colspan=2>PBT<br>Company</td>"	

			sSqlT1 = sSqlT1 & "</tr>"

			ccompcode=""
			MTDRevenueFlash=0
			PBTOPSFlash=0
			INVIncomeFlash=0
			PBTCompanyFlash=0

			MTDRevenueBudget=0
			PBTOPSBudget=0
			INVIncomeBudget=0
			PBTCompanyBudget=0

			Do while not rRs.EOF
				'	
				If ccompcode <> rRs("ccompnamesn2") Then
					If ccompcode<>"" Then
						If textSize="SMALL" then
							sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalSMALL>"	
						Else
							sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalPage8>"	
						End if
					
						
						sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='left' nowrap>" & ccompcode & "</Td>"
						If MTDRevenueFlash <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(MTDRevenueFlash,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End if

						If PBTOPSFlash <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' >" & formatnumber(PBTOPSFlash,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right' >&nbsp;</Td>"
						End if


						If INVIncomeFlash <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeFlash,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End if

						If PBTCompanyFlash <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyFlash,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"	
						End if


						If MTDRevenueBudget <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>" & formatnumber(MTDRevenueBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>&nbsp;</Td>"
						End if

						If PBTOPSBudget <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End if

						If INVIncomeBudget <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End if


						If PBTCompanyBudget <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End if



						If (MTDRevenueFlash - MTDRevenueBudget) < 0 Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/down.png' width='35%'></Td>"
						ElseIf (MTDRevenueFlash - MTDRevenueBudget) > 0 Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/up.png' width='40%'></Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/zero.png' width='60%'></Td>"
						End if						

						If ((MTDRevenueFlash - MTDRevenueBudget) <> 0) then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(MTDRevenueFlash - MTDRevenueBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
						End if
						

						If ((PBTOPSFlash - PBTOPSBudget) < 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
						ElseIf (PBTOPSFlash - PBTOPSBudget) > 0 Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
						End If


						If ((PBTOPSFlash - PBTOPSBudget) <> 0) then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(PBTOPSFlash - PBTOPSBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
						End if						

						
						If ((INVIncomeFlash - INVIncomeBudget) < 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
						ElseIf (INVIncomeFlash - INVIncomeBudget) > 0 Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
						End If
						
						If ((INVIncomeFlash - INVIncomeBudget) <> 0) then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(INVIncomeFlash - INVIncomeBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
						End if


						If ((PBTCompanyFlash - PBTCompanyBudget) < 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
						ElseIf (PBTCompanyFlash - PBTCompanyBudget) > 0 Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
						End If
						
						If ((PBTCompanyFlash - PBTCompanyBudget) <> 0) then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(PBTCompanyFlash - PBTCompanyBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
						End if						



						Totalcol1= Totalcol1 + MTDRevenueFlash
						Totalcol2= Totalcol2 + PBTOPSFlash
						Totalcol3= Totalcol3 + INVIncomeFlash
						Totalcol4= Totalcol4 + PBTCompanyFlash
						Totalcol5= Totalcol5 + MTDRevenueBudget
						Totalcol6= Totalcol6 + PBTOPSBudget
						Totalcol7= Totalcol7 + INVIncomeBudget
						Totalcol8= Totalcol8 + PBTCompanyBudget
						Totalcol9= Totalcol9 + (MTDRevenueFlash - MTDRevenueBudget)
						Totalcol10= Totalcol10 + (PBTOPSFlash - PBTOPSBudget)
						Totalcol11= Totalcol11 + (INVIncomeFlash - INVIncomeBudget)
						Totalcol12= Totalcol12 + (PBTCompanyFlash - PBTCompanyBudget)
						sSqlT1 = sSqlT1 & "</TR>"	

						ccompcode = rRs("ccompnamesn2")
					Else
						'
					End If 
				End If 
				
				ccompcode = rRs("ccompnamesn2")

				If rRs("DESCCD")= "Q5" then
					MTDRevenueFlash = rRs("Flash")
					MTDRevenueBudget= rRs("Budget")
				End If
				
				If rRs("DESCCD")= "R5" then
					PBTOPSFlash = rRs("Flash")
					PBTOPSBudget= rRs("Budget")
				End if

				If rRs("DESCCD")= "E5" then
					INVIncomeFlash = rRs("Flash")
					INVIncomeBudget= rRs("Budget")
				End if

				If rRs("DESCCD")= "K5" then
					PBTCompanyFlash = rRs("Flash")
					PBTCompanyBudget= rRs("Budget")
				End if


				rRs.MoveNext
				

			Loop
		
			'
			If textSize="SMALL" then
				sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalSMALL>"	
			Else
				sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalPage8>"	
			End if
			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='left' nowrap>" & ccompcode & "</Td>"
			If MTDRevenueFlash <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(MTDRevenueFlash,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If PBTOPSFlash <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSFlash,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If INVIncomeFlash <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeFlash,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If PBTCompanyFlash <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyFlash,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 


			If MTDRevenueBudget <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>" & formatnumber(MTDRevenueBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>&nbsp;</Td>"
			End If 

			If PBTOPSBudget <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If INVIncomeBudget <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If PBTCompanyBudget <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 


			

			If ((MTDRevenueFlash - MTDRevenueBudget) < 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/down.png' width='35%'></Td>"
			ElseIf ((MTDRevenueFlash - MTDRevenueBudget) > 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/zero.png' width='60%'></Td>"
			End if						

			If ((MTDRevenueFlash - MTDRevenueBudget) <> 0) then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(MTDRevenueFlash - MTDRevenueBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 



			If ((PBTOPSFlash - PBTOPSBudget) < 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf ((PBTOPSFlash - PBTOPSBudget) > 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End If

			If ((PBTOPSFlash - PBTOPSBudget) <> 0) then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(PBTOPSFlash - PBTOPSBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 
			

			
			If ((INVIncomeFlash - INVIncomeBudget) < 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf ((INVIncomeFlash - INVIncomeBudget) > 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End if

			If ((INVIncomeFlash - INVIncomeBudget) <> 0) then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(INVIncomeFlash - INVIncomeBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 



			If ((PBTCompanyFlash - PBTCompanyBudget) < 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf ((PBTCompanyFlash - PBTCompanyBudget) > 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End If
			
			If ((PBTCompanyFlash - PBTCompanyBudget) <> 0) then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(PBTCompanyFlash - PBTCompanyBudget,0,,(-1)) & "</Td></TR>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td></TR>"
			End If 


			Totalcol1= Totalcol1 + MTDRevenueFlash
			Totalcol2= Totalcol2 + PBTOPSFlash
			Totalcol3= Totalcol3 + INVIncomeFlash
			Totalcol4= Totalcol4 + PBTCompanyFlash
			Totalcol5= Totalcol5 + MTDRevenueBudget
			Totalcol6= Totalcol6 + PBTOPSBudget
			Totalcol7= Totalcol7 + INVIncomeBudget
			Totalcol8= Totalcol8 + PBTCompanyBudget
			Totalcol9= Totalcol9 + (MTDRevenueFlash - MTDRevenueBudget)
			Totalcol10= Totalcol10 + (PBTOPSFlash - PBTOPSBudget)
			Totalcol11= Totalcol11 + (INVIncomeFlash - INVIncomeBudget)
			Totalcol12= Totalcol12 + (PBTCompanyFlash - PBTCompanyBudget)

			If textSize="SMALL" Then
				'sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
				sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleCalculatedFiledSMALL>"	
			ELSE
				'sSqlT1 = sSqlT1 & "<TR Class=RepfontStyle>"	
				sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleCalculatedFiledPage8>"	
			End If
			'   style='padding: 0 0 0 1px;'
			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='left'>Total</Td>"
			If TotalCol1 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 1px 0 0;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol1,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 1px 0 0;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If TotalCol2 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol2,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If TotalCol3 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol3,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If TotalCol4 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol4,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			
			If TotalCol5 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>" & formatnumber(TotalCol5,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>&nbsp;</Td>"
			End If 

			If TotalCol6 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol6,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If TotalCol7 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol7,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If TotalCol8 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol8,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol7,0,,(-1)) & "</Td>"
			'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol8,0,,(-1)) & "</Td>"


			If (TotalCol9) < 0 Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/down.png' width='35%'></Td>"
			ElseIf (TotalCol9) > 0 Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/zero.png' width='60%'></Td>"
			End if						

			If TotalCol9 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol9,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 



			If (TotalCol10) < 0 Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf (TotalCol10) > 0 Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End if						

			If TotalCol10 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol10,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 

	'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol10,0,,(-1)) & "</Td>"

			If (TotalCol11) < 0 Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf (TotalCol11) > 0 Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End if						

			If TotalCol11 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol11,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 

	'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol11,0,,(-1)) & "</Td>"

			If (TotalCol12) < 0 Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf (TotalCol12) > 0 Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End if						

			If TotalCol12 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(TotalCol12,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
			End If 

	'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(TotalCol12,0,,(-1)) & "</Td></TR>"
			

'			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >&nbsp;</td>"
'			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' align='right'>" & formatnumber(TotalCol9,0,,(-1)) & "</Td>"
'			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >&nbsp;</td>"
'			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' align='right'>" & formatnumber(TotalCol10,0,,(-1)) & "</Td>"
'			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >&nbsp;</td>"
'			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' align='right'>" & formatnumber(TotalCol11,0,,(-1)) & "</Td>"
'			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >&nbsp;</td>"
'			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' align='right'>" & formatnumber(TotalCol12,0,,(-1)) & "</Td>"

		
			sSqlT1 = sSqlT1 & "</table>"	
		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (I)**********</Td>"
		End if		
		rRs.close



		sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
		sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><TD >" 
		sSql = sSql & "Amount in RO '000"
		sSql = sSql & "</Td >" 
		sSql = sSql & "</TR>" 
		If textSize<>"SMALL" Then
			sSql = sSql & "<TR Class=GreenTransparentHeadingNormal><TD  >" 
		Else
			sSql = sSql & "<TR Class=GreenTransparentHeadingSmall><TD  >" 
		End If 
		sSql = sSql & Title1
		sSql = sSql & "</Td >" 
		sSql = sSql & "</TR>" 
		sSql = sSql & "</Table>"
		response.write sSql	

		Response.write sSqlT1 
		'Response.write "<br>" 
		'Response.write "<br>" 
		If textSize<>"SMALL" Then
			Response.write "</TD></TR><TR Class=RepContentFontSMALL><TD>&nbsp;</TD></TR><TR RepContentFontSMALL><TD  valign='top'>" 
		End if

		sSqlT1=""

	'	response.write sSql2
		rRs.Open Trim(sSql2),cCn

		SRLNO = 0
		PageNumber=0
		Line_Count=1
		CompID=""
		CompanyName=""
		SectorNo=""

		Totalcol1=0
		Totalcol2=0
		Totalcol3=0
		Totalcol4=0
		Totalcol5=0
		Totalcol6=0
		Totalcol7=0
		Totalcol8=0
		Totalcol9=0
		Totalcol10=0
		Totalcol11=0
		Totalcol12=0

		sSql = ""
		sSqlT1= ""

		If Not rRs.Eof Then
			'

			If textSize="SMALL" Then
				sSqlT1 = sSqlT1 & "<Table border=0  width='100%' cellPadding=0 cellspacing=0>"
			Else
				sSqlT1 = sSqlT1 & "<Table border=0  width='100%' cellPadding=1 cellspacing=0>"
			End If
			
			'If textSize="SMALL" Then
			'	sSqlT1 = sSqlT1 & "<TR Class=AllReportHeaderSMALL><TD class='no-top-right-border' Align='center' colspan=13 nowrap>" & Title2 & "</td><TD class='no-top-right-border' colspan=4 align=right>Amount (RO '000)</td></TR>"	
			'Else
		    '	sSqlT1 = sSqlT1 & "<TR Class=AllReportHeader><TD class='no-top-right-border' Align='center' colspan=13 nowrap>" & Title2 & "</td><TD class='no-top-right-border' colspan=4 align=right>Amount (RO '000)</td></TR>"	
			'End if

			If period = "M" Then
				If textSize="SMALL" Then
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
				Else
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStylePage8>"	
				End if
				sSqlT1 = sSqlT1 & "<TD class='no-right-border' Align='center' rowspan=2 >Company</td>"	
				sSqlT1 = sSqlT1 & "<TD class='no-right-border' Align='center' colspan=4>MTD Flash " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
				sSqlT1 = sSqlT1 & "<TD class='no-right-border_thick' Align='center' colspan=4>MTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & CInt(TheReportYear)-1 & "</td>"	
				sSqlT1 = sSqlT1 & "<TD class='border-ALL_left_thick' Align='center' colspan=8>MTD Flash vs LY</td>"	
				sSqlT1 = sSqlT1 & "</tr>"
			elseIf period = "Y" Then
				If textSize="SMALL" Then
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
				Else
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStylePage8>"	
				End if

				sSqlT1 = sSqlT1 & "<TD class='no-right-border' Align='center' rowspan=2 >Company</td>"	
				sSqlT1 = sSqlT1 & "<TD class='no-right-border' Align='center' colspan=4>YTD Flash " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & "</td>"	
				sSqlT1 = sSqlT1 & "<TD class='no-right-border_thick' Align='center' colspan=4>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & CInt(TheReportYear)-1 & "</td>"	
				sSqlT1 = sSqlT1 & "<TD class='border-ALL_left_thick' Align='center' colspan=8>YTD Flash vs LY</td>"	
				sSqlT1 = sSqlT1 & "</tr>"
			End If 
				If textSize="SMALL" Then
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
				Else
					sSqlT1 = sSqlT1 & "<TR Class=RepfontStylePage8>"	
				End if

			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >Revenue</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' nowrap >PBT Ops</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >Invst<br>Income</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >PBT<br>Company</td>"	

			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border_thick' Align='center' >Revenue</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' nowrap >PBT Ops</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >Invst<br>Income</td>"	
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >PBT<br>Company</td>"	

			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >&nbsp;</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-right-left-top-border' Align='center' >Revenue</td>"
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >&nbsp;</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-right-left-top-border' Align='center' nowrap >PBT Ops</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >&nbsp;</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-right-left-top-border' Align='center' >Invst<br>Income</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' >&nbsp;</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-left-border' Align='center' >PBT<br>Company</td>"	

			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center'></Td>"
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border_thick' Align='center' colspan=2 >Revenue</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' ></Td>"
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' nowrap colspan=2>PBT Ops</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' ></Td>"
			sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' colspan=2>Invst<br>Income</td>"	
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' Align='center' ></Td>"
			sSqlT1 = sSqlT1 & "<TD class='no-top-border' Align='center' colspan=2>PBT<br>Company</td>"	

			sSqlT1 = sSqlT1 & "</tr>"

			ccompcode=""
			MTDRevenueFlash=0
			PBTOPSFlash=0
			INVIncomeFlash=0
			PBTCompanyFlash=0

			MTDRevenueBudget=0
			PBTOPSBudget=0
			INVIncomeBudget=0
			PBTCompanyBudget=0

			Do while not rRs.EOF
				'	
				If ccompcode <> rRs("ccompnamesn2") Then
					If ccompcode<>"" then
						If textSize="SMALL" Then
							sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalSMALL>"		
						Else
							sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalPage8>"	
						End if

						
						sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='left' nowrap>" & ccompcode & "</Td>"

						If MTDRevenueFlash <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(MTDRevenueFlash,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End If 

						If PBTOPSFlash <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSFlash,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End If 

						If INVIncomeFlash <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeFlash,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End If 

						If PBTCompanyFlash <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyFlash,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End If 


						If MTDRevenueBudget <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>" & formatnumber(MTDRevenueBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>&nbsp;</Td>"
						End If 

						If PBTOPSBudget <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End If 

						If INVIncomeBudget <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End If 

						If PBTCompanyBudget <> 0 then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
						End If 

'						sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeBudget,0,,(-1)) & "</Td>"
'						sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyBudget,0,,(-1)) & "</Td>"

						If ((MTDRevenueFlash - MTDRevenueBudget) < 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/down.png' width='35%'></Td>"
						ElseIf ((MTDRevenueFlash - MTDRevenueBudget) > 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/up.png' width='40%'></Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/zero.png' width='60%'></Td>"
						End if						

						If ((MTDRevenueFlash - MTDRevenueBudget) <> 0) then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(MTDRevenueFlash - MTDRevenueBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
						End If 

						'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(MTDRevenueFlash - MTDRevenueBudget,0,,(-1)) & "</Td>"

						If (PBTOPSFlash - PBTOPSBudget) < 0 Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
						ElseIf (PBTOPSFlash - PBTOPSBudget) > 0 Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
						End If

						If ((PBTOPSFlash - PBTOPSBudget) <> 0) then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(PBTOPSFlash - PBTOPSBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
						End If 
'1						
							'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(PBTOPSFlash - PBTOPSBudget,0,,(-1)) & "</Td>"
						
						If ((INVIncomeFlash - INVIncomeBudget) < 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
						ElseIf ((INVIncomeFlash - INVIncomeBudget) > 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
						End if
						If ((INVIncomeFlash - INVIncomeBudget) <> 0) then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(INVIncomeFlash - INVIncomeBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
						End If 

						'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(INVIncomeFlash - INVIncomeBudget,0,,(-1)) & "</Td>"

						If ((PBTCompanyFlash - PBTCompanyBudget) < 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
						ElseIf ((PBTCompanyFlash - PBTCompanyBudget) > 0) Then
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
						End If
						If ((PBTCompanyFlash - PBTCompanyBudget) <> 0) then
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(PBTCompanyFlash - PBTCompanyBudget,0,,(-1)) & "</Td>"
						Else
							sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
						End If 
						
						'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(PBTCompanyFlash - PBTCompanyBudget,0,,(-1)) & "</Td>"

						

						Totalcol1= Totalcol1 + MTDRevenueFlash
						Totalcol2= Totalcol2 + PBTOPSFlash
						Totalcol3= Totalcol3 + INVIncomeFlash
						Totalcol4= Totalcol4 + PBTCompanyFlash
						Totalcol5= Totalcol5 + MTDRevenueBudget
						Totalcol6= Totalcol6 + PBTOPSBudget
						Totalcol7= Totalcol7 + INVIncomeBudget
						Totalcol8= Totalcol8 + PBTCompanyBudget
						Totalcol9= Totalcol9 + (MTDRevenueFlash - MTDRevenueBudget)
						Totalcol10= Totalcol10 + (PBTOPSFlash - PBTOPSBudget)
						Totalcol11= Totalcol11 + (INVIncomeFlash - INVIncomeBudget)
						Totalcol12= Totalcol12 + (PBTCompanyFlash - PBTCompanyBudget)

						
						sSqlT1 = sSqlT1 & "</TR>"	
						ccompcode = rRs("ccompnamesn2")
					Else
						'
					End If 
				End If 
				
				ccompcode = rRs("ccompnamesn2")

				If rRs("DESCCD")= "Q5" then
					MTDRevenueFlash = rRs("Flash")
					MTDRevenueBudget= rRs("Budget")
				End If
				
				If rRs("DESCCD")= "R5" then
					PBTOPSFlash = rRs("Flash")
					PBTOPSBudget= rRs("Budget")
				End if

				If rRs("DESCCD")= "E5" then
					INVIncomeFlash = rRs("Flash")
					INVIncomeBudget= rRs("Budget")
				End if

				If rRs("DESCCD")= "K5" then
					PBTCompanyFlash = rRs("Flash")
					PBTCompanyBudget= rRs("Budget")
				End if


				rRs.MoveNext
				

			Loop

			'
			If textSize="SMALL" Then
				sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalSMALL>"	
			Else
				sSqlT1 = sSqlT1 & "<TR Class=RepContentFontNormalPage8>"	
			End if
			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='left' nowrap>" & ccompcode & "</Td>"

			If MTDRevenueFlash <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(MTDRevenueFlash,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If PBTOPSFlash <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSFlash,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If INVIncomeFlash <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeFlash,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If PBTCompanyFlash <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyFlash,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSFlash,0,,(-1)) & "</Td>"
			'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeFlash,0,,(-1)) & "</Td>"
			'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyFlash,0,,(-1)) & "</Td>"

			If MTDRevenueBudget <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>" & formatnumber(MTDRevenueBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>&nbsp;</Td>"
			End If 


			If PBTOPSBudget <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If INVIncomeBudget <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			If PBTCompanyBudget <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If 

			'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTOPSBudget,0,,(-1)) & "</Td>"
			'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(INVIncomeBudget,0,,(-1)) & "</Td>"
			'sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyBudget,0,,(-1)) & "</Td>"
			

			If ((MTDRevenueFlash - MTDRevenueBudget) < 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/down.png' width='35%'></Td>"
			ElseIf ((MTDRevenueFlash - MTDRevenueBudget) > 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/zero.png' width='60%'></Td>"
			End if						

			If ((MTDRevenueFlash - MTDRevenueBudget) <> 0) then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(MTDRevenueFlash - MTDRevenueBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 



			If ((PBTOPSFlash - PBTOPSBudget) < 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf ((PBTOPSFlash - PBTOPSBudget) > 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End If

			If ((PBTOPSFlash - PBTOPSBudget) <> 0) then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(PBTOPSFlash - PBTOPSBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 
			
'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(PBTOPSFlash - PBTOPSBudget,0,,(-1)) & "</Td>"
			
			If ((INVIncomeFlash - INVIncomeBudget) < 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf ((INVIncomeFlash - INVIncomeBudget) > 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End if
			If ((INVIncomeFlash - INVIncomeBudget) <> 0) then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(INVIncomeFlash - INVIncomeBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If 

'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(INVIncomeFlash - INVIncomeBudget,0,,(-1)) & "</Td>"

			If ((PBTCompanyFlash - PBTCompanyBudget) < 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf ((PBTCompanyFlash - PBTCompanyBudget) > 0) Then
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End If

			If ((PBTCompanyFlash - PBTCompanyBudget) <> 0) then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(PBTCompanyFlash - PBTCompanyBudget,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
			End If 

'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(PBTCompanyFlash - PBTCompanyBudget,0,,(-1)) & "</Td></TR>"

			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' align='right'>" & formatnumber(MTDRevenueFlash - MTDRevenueBudget,0,,(-1)) & "</Td>"
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' align='right'>" & formatnumber(PBTOPSFlash - PBTOPSBudget,0,,(-1)) & "</Td>"
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' align='right'>" & formatnumber(INVIncomeFlash - INVIncomeBudget,0,,(-1)) & "</Td>"
			'sSqlT1 = sSqlT1 & "<TD class='no-top-right-border' align='right'>" & formatnumber(PBTCompanyFlash - PBTCompanyBudget,0,,(-1)) & "</Td>"

			Totalcol1= Totalcol1 + MTDRevenueFlash
			Totalcol2= Totalcol2 + PBTOPSFlash
			Totalcol3= Totalcol3 + INVIncomeFlash
			Totalcol4= Totalcol4 + PBTCompanyFlash
			Totalcol5= Totalcol5 + MTDRevenueBudget
			Totalcol6= Totalcol6 + PBTOPSBudget
			Totalcol7= Totalcol7 + INVIncomeBudget
			Totalcol8= Totalcol8 + PBTCompanyBudget
			Totalcol9= Totalcol9 + (MTDRevenueFlash - MTDRevenueBudget)
			Totalcol10= Totalcol10 + (PBTOPSFlash - PBTOPSBudget)
			Totalcol11= Totalcol11 + (INVIncomeFlash - INVIncomeBudget)
			Totalcol12= Totalcol12 + (PBTCompanyFlash - PBTCompanyBudget)


			'If textSize="SMALL" Then
			'	sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
			'Else
			'	sSqlT1 = sSqlT1 & "<TR Class=RepfontStyle>"	
			'End If

			If textSize="SMALL" Then
				'sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleSMALL>"	
				sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleCalculatedFiledSMALL>"	
			ELSE
				'sSqlT1 = sSqlT1 & "<TR Class=RepfontStyle>"	
				sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleCalculatedFiledPage8>"	
			End IF
			
			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='left'>Total</Td>"
			If TotalCol1 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol1,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If
			If TotalCol2 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol2,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If
			If TotalCol3 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol3,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If
			If TotalCol4 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol4,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End if			
'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol2,0,,(-1)) & "</Td>"
'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol3,0,,(-1)) & "</Td>"
'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol4,0,,(-1)) & "</Td>"


			If TotalCol5 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>" & formatnumber(TotalCol5,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border_thick' align='right'>&nbsp;</Td>"
			End If

			

			If TotalCol6 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol6,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If
			If TotalCol7 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol7,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End If
			If TotalCol8 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol8,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>&nbsp;</Td>"
			End if			
'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol6,0,,(-1)) & "</Td>"
'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol7,0,,(-1)) & "</Td>"
'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border' align='right'>" & formatnumber(TotalCol8,0,,(-1)) & "</Td>"

			If (TotalCol9) < 0 Then
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/down.png' width='35%'></Td>"
			ElseIf (TotalCol9) > 0 Then
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border_thick'><img src='images/zero.png' width='60%'></Td>"
			End if						

			If TotalCol9 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol9,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If
			
'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol9,0,,(-1)) & "</Td>"

			If (TotalCol10) < 0 Then
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf (TotalCol10) > 0 Then
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End if						

			If TotalCol10 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol10,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If

'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol10,0,,(-1)) & "</Td>"

			If (TotalCol11) < 0 Then
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf (TotalCol11) > 0 Then
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End if						
			If TotalCol11 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol11,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>&nbsp;</Td>"
			End If

'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-right-left-top-border' align='right'>" & formatnumber(TotalCol11,0,,(-1)) & "</Td>"

			If (TotalCol12) < 0 Then
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' width='35%'></Td>"
			ElseIf (TotalCol12) > 0 Then
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' width='40%'></Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD  Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' width='60%'></Td>"
			End if						
			If TotalCol12 <> 0 then
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(TotalCol12,0,,(-1)) & "</Td>"
			Else
				sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>&nbsp;</Td>"
			End If

'			sSqlT1 = sSqlT1 & "<TD style='padding: 0 2px 0 2px;' class='no-top-left-border' align='right'>" & formatnumber(TotalCol12,0,,(-1)) & "</Td></TR>"
			
			If textSize="SMALL" Then
				sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleReportRemarkSMALL>"
			Else
				sSqlT1 = sSqlT1 & "<TR Class=RepfontStyleReportRemark>"
			End If
			

			sSqlT1 = sSqlT1 & "<TD  colspan=17>*Revenue includes Sales, Commission and Other Income."
			sSqlT1 = sSqlT1 & "</TD>"
			sSqlT1 = sSqlT1 & "</TR>"	

			sSqlT1 = sSqlT1 & "</table>"	
		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched (J)**********</Td>"
		End if		

rRs.close
		sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
		If textSize<>"SMALL" Then
			sSql = sSql & "<TR Class=GreenTransparentHeadingNormal><TD  >" 
		Else
			sSql = sSql & "<TR Class=GreenTransparentHeadingSmall><TD >" 
		End If 
		sSql = sSql & Title2
		sSql = sSql & "</Td >" 
		sSql = sSql & "</TR>" 
		sSql = sSql & "</Table>"
		response.write sSql
		Response.write sSqlT1 

	End function

	Private Function Page17()

		sSql="" 
		sSql= "SELECT 'Manufacturing' Sector, DESCCD, DESCRIPTION,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_MTD_BAL,'R5', ACT_MTD_BAL,'M5',ACT_MTD_BAL,'E5',ACT_MTD_BAL,0)) as ACT_MTD_BAL , " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', BGT_MTD_BAL,'R5', BGT_MTD_BAL,'M5',BGT_MTD_BAL,'E5',BGT_MTD_BAL,0)) as BGT_MTD_BAL , " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_MTD_BAL_PY1,'R5', ACT_MTD_BAL_PY1,'M5',ACT_MTD_BAL_PY1,'E5',ACT_MTD_BAL_PY1,0)) as ACT_MTD_BAL_PY1, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL,'R5', ACT_YTD_BAL,'M5',ACT_YTD_BAL,'E5',ACT_YTD_BAL,0)) as ACT_YTD_BAL ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', BGT_YTD_BAL,'R5', BGT_YTD_BAL,'M5',BGT_YTD_BAL,'E5',BGT_YTD_BAL,0)) as BGT_YTD_BAL ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY1,'R5', ACT_YTD_BAL_PY1,'M5',ACT_YTD_BAL_PY1,'E5',ACT_YTD_BAL_PY1,0)) as ACT_YTD_BAL_PY1 ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY2,'R5', ACT_YTD_BAL_PY2,'M5',ACT_YTD_BAL_PY2,'E5',ACT_YTD_BAL_PY2,0)) as ACT_YTD_BAL_PY2, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY3,'R5', ACT_YTD_BAL_PY3,'M5',ACT_YTD_BAL_PY3,'E5',ACT_YTD_BAL_PY3,0)) as ACT_YTD_BAL_PY3, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY4,'R5', ACT_YTD_BAL_PY4,'M5',ACT_YTD_BAL_PY4,'E5',ACT_YTD_BAL_PY4,0)) as ACT_YTD_BAL_PY4 " 
		sSql= sSql & " FROM MISTXN LEFT JOIN COMPANY ON MISTXN.CCOMPCODE = COMPANY.CCOMPCODE  " 
		sSql= sSql & " where Trim(cMonth)='" & TheReportMonth & "'  " 
		sSql= sSql & " AND Trim(cYear)='" & TheReportYear & "'  " 
		sSql= sSql & " AND DESCCD IN ('Q5','R5','M5','E5') " 
		sSql= sSql & " AND COMPANY.CATEGORY = 'Manufacturing'  " 
		sSql= sSql & " AND COMPANY.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') " 
		sSql= sSql & " Group By cMonth,cYear, DESCCD, DESCRIPTION " 
		sSql= sSql & " Union  " 
		sSql= sSql & " SELECT 'Service And Investment' , DESCCD, DESCRIPTION,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_MTD_BAL,'R5', ACT_MTD_BAL,'M5',ACT_MTD_BAL,'E5',ACT_MTD_BAL,0)) as ACT_MTD_BAL , " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', BGT_MTD_BAL,'R5', BGT_MTD_BAL,'M5',BGT_MTD_BAL,'E5',BGT_MTD_BAL,0)) as BGT_MTD_BAL , " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_MTD_BAL_PY1,'R5', ACT_MTD_BAL_PY1,'M5',ACT_MTD_BAL_PY1,'E5',ACT_MTD_BAL_PY1,0)) as ACT_MTD_BAL_PY1, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL,'R5', ACT_YTD_BAL,'M5',ACT_YTD_BAL,'E5',ACT_YTD_BAL,0)) as ACT_YTD_BAL ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', BGT_YTD_BAL,'R5', BGT_YTD_BAL,'M5',BGT_YTD_BAL,'E5',BGT_YTD_BAL,0)) as BGT_YTD_BAL ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY1,'R5', ACT_YTD_BAL_PY1,'M5',ACT_YTD_BAL_PY1,'E5',ACT_YTD_BAL_PY1,0)) as ACT_YTD_BAL_PY1 ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY2,'R5', ACT_YTD_BAL_PY2,'M5',ACT_YTD_BAL_PY2,'E5',ACT_YTD_BAL_PY2,0)) as ACT_YTD_BAL_PY2, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY3,'R5', ACT_YTD_BAL_PY3,'M5',ACT_YTD_BAL_PY3,'E5',ACT_YTD_BAL_PY3,0)) as ACT_YTD_BAL_PY3, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY4,'R5', ACT_YTD_BAL_PY4,'M5',ACT_YTD_BAL_PY4,'E5',ACT_YTD_BAL_PY4,0)) as ACT_YTD_BAL_PY4 " 
		sSql= sSql & " FROM MISTXN LEFT JOIN COMPANY ON MISTXN.CCOMPCODE = COMPANY.CCOMPCODE  " 
		sSql= sSql & " where Trim(cMonth)='" & TheReportMonth & "'  " 
		sSql= sSql & " AND Trim(cYear)='" & TheReportYear & "'  " 
		sSql= sSql & " AND DESCCD IN ('Q5','R5','M5','E5') " 
		sSql= sSql & " AND COMPANY.CATEGORY like ('Service%')  " 
		sSql= sSql & " AND COMPANY.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') " 
		sSql= sSql & " Group By cMonth,cYear, DESCCD, DESCRIPTION " 
		sSql= sSql & " Union All "
		sSql= sSql & " SELECT 'Trading And Contracting' Sector, DESCCD, DESCRIPTION,  "
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_MTD_BAL,'R5', ACT_MTD_BAL,'M5',ACT_MTD_BAL,'E5',ACT_MTD_BAL,0)) as ACT_MTD_BAL , " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', BGT_MTD_BAL,'R5', BGT_MTD_BAL,'M5',BGT_MTD_BAL,'E5',BGT_MTD_BAL,0)) as BGT_MTD_BAL , " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_MTD_BAL_PY1,'R5', ACT_MTD_BAL_PY1,'M5',ACT_MTD_BAL_PY1,'E5',ACT_MTD_BAL_PY1,0)) as ACT_MTD_BAL_PY1, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL,'R5', ACT_YTD_BAL,'M5',ACT_YTD_BAL,'E5',ACT_YTD_BAL,0)) as ACT_YTD_BAL ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', BGT_YTD_BAL,'R5', BGT_YTD_BAL,'M5',BGT_YTD_BAL,'E5',BGT_YTD_BAL,0)) as BGT_YTD_BAL ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY1,'R5', ACT_YTD_BAL_PY1,'M5',ACT_YTD_BAL_PY1,'E5',ACT_YTD_BAL_PY1,0)) as ACT_YTD_BAL_PY1 ,  " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY2,'R5', ACT_YTD_BAL_PY2,'M5',ACT_YTD_BAL_PY2,'E5',ACT_YTD_BAL_PY2,0)) as ACT_YTD_BAL_PY2, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY3,'R5', ACT_YTD_BAL_PY3,'M5',ACT_YTD_BAL_PY3,'E5',ACT_YTD_BAL_PY3,0)) as ACT_YTD_BAL_PY3, " 
		sSql= sSql & " sum(DECODE(DESCCD,'Q5', ACT_YTD_BAL_PY4,'R5', ACT_YTD_BAL_PY4,'M5',ACT_YTD_BAL_PY4,'E5',ACT_YTD_BAL_PY4,0)) as ACT_YTD_BAL_PY4 " 
		sSql= sSql & " FROM MISTXN LEFT JOIN COMPANY ON MISTXN.CCOMPCODE = COMPANY.CCOMPCODE  "
		sSql= sSql & " where Trim(cMonth)='" & TheReportMonth & "'  "
		sSql= sSql & " AND Trim(cYear)='" & TheReportYear & "'  "
		sSql= sSql & " AND DESCCD IN ('Q5','R5','M5','E5') "
		sSql= sSql & " AND COMPANY.CATEGORY in ('Trading','Contracting')  "
		sSql= sSql & " AND COMPANY.CCOMPCODE NOT IN ('OMHC','OMC','OFCC') "
		sSql= sSql & " Group By cMonth,cYear, DESCCD, DESCRIPTION "
		sSql= sSql & " order by 1 "

'response.write sSql
'response.end
		rRs.Open Trim(sSql),cCn

		RevenueManufacturingCol1=0
		RevenueManufacturingCol2=0
		RevenueManufacturingCol3=0
		RevenueManufacturingCol4=0
		RevenueManufacturingCol5=0
		RevenueManufacturingCol6=0
		RevenueManufacturingCol7=0
		RevenueManufacturingCol8=0
		RevenueManufacturingCol9=0

		RevenueTradingAndContractingCol1=0
		RevenueTradingAndContractingCol2=0
		RevenueTradingAndContractingCol3=0
		RevenueTradingAndContractingCol4=0
		RevenueTradingAndContractingCol5=0
		RevenueTradingAndContractingCol6=0
		RevenueTradingAndContractingCol7=0
		RevenueTradingAndContractingCol8=0
		RevenueTradingAndContractingCol9=0

		RevenueServiceAndInvestmentCol1=0
		RevenueServiceAndInvestmentCol2=0
		RevenueServiceAndInvestmentCol3=0
		RevenueServiceAndInvestmentCol4=0
		RevenueServiceAndInvestmentCol5=0
		RevenueServiceAndInvestmentCol6=0
		RevenueServiceAndInvestmentCol7=0
		RevenueServiceAndInvestmentCol8=0
		RevenueServiceAndInvestmentCol9=0

		'-*-
		PBTOperationManufacturingCol1=0
		PBTOperationManufacturingCol2=0
		PBTOperationManufacturingCol3=0
		PBTOperationManufacturingCol4=0
		PBTOperationManufacturingCol5=0
		PBTOperationManufacturingCol6=0
		PBTOperationManufacturingCol7=0
		PBTOperationManufacturingCol8=0
		PBTOperationManufacturingCol9=0

		PBTOperationTradingAndContractingCol1=0
		PBTOperationTradingAndContractingCol2=0
		PBTOperationTradingAndContractingCol3=0
		PBTOperationTradingAndContractingCol4=0
		PBTOperationTradingAndContractingCol5=0
		PBTOperationTradingAndContractingCol6=0
		PBTOperationTradingAndContractingCol7=0
		PBTOperationTradingAndContractingCol8=0
		PBTOperationTradingAndContractingCol9=0

		PBTOperationServiceAndInvestmentCol1=0
		PBTOperationServiceAndInvestmentCol2=0
		PBTOperationServiceAndInvestmentCol3=0
		PBTOperationServiceAndInvestmentCol4=0
		PBTOperationServiceAndInvestmentCol5=0
		PBTOperationServiceAndInvestmentCol6=0
		PBTOperationServiceAndInvestmentCol7=0
		PBTOperationServiceAndInvestmentCol8=0
		PBTOperationServiceAndInvestmentCol9=0
		'-*-
		'-*-
		M5ManufacturingCol1=0
		M5ManufacturingCol2=0
		M5ManufacturingCol3=0
		M5ManufacturingCol4=0
		M5ManufacturingCol5=0
		M5ManufacturingCol6=0
		M5ManufacturingCol7=0
		M5ManufacturingCol8=0
		M5ManufacturingCol9=0

		M5TradingAndContractingCol1=0
		M5TradingAndContractingCol2=0
		M5TradingAndContractingCol3=0
		M5TradingAndContractingCol4=0
		M5TradingAndContractingCol5=0
		M5TradingAndContractingCol6=0
		M5TradingAndContractingCol7=0
		M5TradingAndContractingCol8=0
		M5TradingAndContractingCol9=0

		M5ServiceAndInvestmentCol1=0
		M5ServiceAndInvestmentCol2=0
		M5ServiceAndInvestmentCol3=0
		M5ServiceAndInvestmentCol4=0
		M5ServiceAndInvestmentCol5=0
		M5ServiceAndInvestmentCol6=0
		M5ServiceAndInvestmentCol7=0
		M5ServiceAndInvestmentCol8=0
		M5ServiceAndInvestmentCol9=0
		'-*-
		E5ManufacturingCol1=0
		E5ManufacturingCol2=0
		E5ManufacturingCol3=0
		E5ManufacturingCol4=0
		E5ManufacturingCol5=0
		E5ManufacturingCol6=0
		E5ManufacturingCol7=0
		E5ManufacturingCol8=0
		E5ManufacturingCol9=0

		E5TradingAndContractingCol1=0
		E5TradingAndContractingCol2=0
		E5TradingAndContractingCol3=0
		E5TradingAndContractingCol4=0
		E5TradingAndContractingCol5=0
		E5TradingAndContractingCol6=0
		E5TradingAndContractingCol7=0
		E5TradingAndContractingCol8=0
		E5TradingAndContractingCol9=0

		E5ServiceAndInvestmentCol1=0
		E5ServiceAndInvestmentCol2=0
		E5ServiceAndInvestmentCol3=0
		E5ServiceAndInvestmentCol4=0
		E5ServiceAndInvestmentCol5=0
		E5ServiceAndInvestmentCol6=0
		E5ServiceAndInvestmentCol7=0
		E5ServiceAndInvestmentCol8=0
		E5ServiceAndInvestmentCol9=0

		If Not rRs.Eof Then
			'
			sSql = ""

	sSql =  "<Table border=0  width='100%' cellPadding=1 cellspacing=0	>"
	sSql = sSql & "<TR Class=RepfontStyleCommonTransparentHeading><Td >" 
	sSql = sSql & "Amount in RO '000"
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "<TR Class=GreenTransparentHeading><Td >" 
	sSql = sSql & "Sectorwise Summary of Group Results" 
	sSql = sSql & "</Td >" 
	sSql = sSql & "</TR>" 
	sSql = sSql & "</Table>"
'	response.write sSql	

			sSql = sSql & "<Table border=0  width='100%' cellPadding=3 cellspacing=0>"
			
			'sSql = sSql & "<TR Class=AllReportHeader><Td Align='center' colspan=11>Sectorwise Summary of Group Results for the period " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & " </td><td colspan=3 align=right>Amount (RO '000)</td></TR>"	


			sSql = sSql & "<TR Class=RepfontStylePage17><Td style='padding: 2px 2px 2px 2px;' class='no-right-border' Align='center' valign='center' width='30%' rowspan=2 >Sector</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;'  class='no-right-border_thick' Align='center' colspan=5>MTD " & FullMonthsArr(cint(TheReportMonth)-1) & " </td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-right-border_thick' Align='center' colspan=5>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & " </td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='border-ALL_left_thick' Align='center' colspan=3>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & " </td>"	
			sSql = sSql & "</tr>"


			sSql = sSql & "<TR Class=RepfontStylePage17>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' Align='center' >Flash<br>" & TheReportYear & "</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >Budget<br>" & TheReportYear & "</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >Actual<br>" & CInt(TheReportYear) - 1  & "</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >vs Budget<br>" & TheReportYear & " </td>"
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >vs LY<br>" & CInt(TheReportYear) - 1  & "</td>"	

			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' Align='center' >Flash<br>" & TheReportYear & "</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >Budget<br>" & TheReportYear & "</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >Actual<br>" & CInt(TheReportYear) - 1  & "</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >vs Budget<br>" & TheReportYear & " </td>"
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >vs LY<br>" & CInt(TheReportYear) - 1  & "</td>"	

			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' Align='center' >Actual<br>" & CInt(TheReportYear) - 2  & "</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-right-border' Align='center' >Actual<br>" & CInt(TheReportYear) - 3  & "</td>"	
			sSql = sSql & "<Td style='padding: 2px 2px 2px 2px;' class='no-top-border' Align='center' >Actual<br>" & CInt(TheReportYear) - 4  & "</td>"	

			sSql = sSql & "</tr>"

			Do while not rRs.EOF
				'	
				If rRs("DESCCD") ="Q5" Then
					If rRs("Sector") = "Manufacturing" Then
						RevenueManufacturingCol1=rRs("ACT_MTD_BAL")
						RevenueManufacturingCol2=rRs("BGT_MTD_BAL")
						RevenueManufacturingCol3=rRs("ACT_MTD_BAL_PY1")
						RevenueManufacturingCol4=rRs("ACT_YTD_BAL")
						RevenueManufacturingCol5=rRs("BGT_YTD_BAL")
						RevenueManufacturingCol6=rRs("ACT_YTD_BAL_PY1")
						RevenueManufacturingCol7=rRs("ACT_YTD_BAL_PY2")
						RevenueManufacturingCol8=rRs("ACT_YTD_BAL_PY3")
						RevenueManufacturingCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
					If rRs("Sector") = "Trading And Contracting" Then
						RevenueTradingAndContractingCol1=rRs("ACT_MTD_BAL")
						RevenueTradingAndContractingCol2=rRs("BGT_MTD_BAL")
						RevenueTradingAndContractingCol3=rRs("ACT_MTD_BAL_PY1")
						RevenueTradingAndContractingCol4=rRs("ACT_YTD_BAL")
						RevenueTradingAndContractingCol5=rRs("BGT_YTD_BAL")
						RevenueTradingAndContractingCol6=rRs("ACT_YTD_BAL_PY1")
						RevenueTradingAndContractingCol7=rRs("ACT_YTD_BAL_PY2")
						RevenueTradingAndContractingCol8=rRs("ACT_YTD_BAL_PY3")
						RevenueTradingAndContractingCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
					If rRs("Sector") = "Service And Investment" Then
						RevenueServiceAndInvestmentCol1=rRs("ACT_MTD_BAL")
						RevenueServiceAndInvestmentCol2=rRs("BGT_MTD_BAL")
						RevenueServiceAndInvestmentCol3=rRs("ACT_MTD_BAL_PY1")
						RevenueServiceAndInvestmentCol4=rRs("ACT_YTD_BAL")
						RevenueServiceAndInvestmentCol5=rRs("BGT_YTD_BAL")
						RevenueServiceAndInvestmentCol6=rRs("ACT_YTD_BAL_PY1")
						RevenueServiceAndInvestmentCol7=rRs("ACT_YTD_BAL_PY2")
						RevenueServiceAndInvestmentCol8=rRs("ACT_YTD_BAL_PY3")
						RevenueServiceAndInvestmentCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
				End if

				If rRs("DESCCD") ="R5" Then
					If rRs("Sector") = "Manufacturing" Then
						PBTOperationManufacturingCol1=rRs("ACT_MTD_BAL")
						PBTOperationManufacturingCol2=rRs("BGT_MTD_BAL")
						PBTOperationManufacturingCol3=rRs("ACT_MTD_BAL_PY1")
						PBTOperationManufacturingCol4=rRs("ACT_YTD_BAL")
						PBTOperationManufacturingCol5=rRs("BGT_YTD_BAL")
						PBTOperationManufacturingCol6=rRs("ACT_YTD_BAL_PY1")
						PBTOperationManufacturingCol7=rRs("ACT_YTD_BAL_PY2")
						PBTOperationManufacturingCol8=rRs("ACT_YTD_BAL_PY3")
						PBTOperationManufacturingCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
					If rRs("Sector") = "Trading And Contracting" Then
						PBTOperationTradingAndContractingCol1=rRs("ACT_MTD_BAL")
						PBTOperationTradingAndContractingCol2=rRs("BGT_MTD_BAL")
						PBTOperationTradingAndContractingCol3=rRs("ACT_MTD_BAL_PY1")
						PBTOperationTradingAndContractingCol4=rRs("ACT_YTD_BAL")
						PBTOperationTradingAndContractingCol5=rRs("BGT_YTD_BAL")
						PBTOperationTradingAndContractingCol6=rRs("ACT_YTD_BAL_PY1")
						PBTOperationTradingAndContractingCol7=rRs("ACT_YTD_BAL_PY2")
						PBTOperationTradingAndContractingCol8=rRs("ACT_YTD_BAL_PY3")
						PBTOperationTradingAndContractingCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
					If rRs("Sector") = "Service And Investment" Then
						PBTOperationServiceAndInvestmentCol1=rRs("ACT_MTD_BAL")
						PBTOperationServiceAndInvestmentCol2=rRs("BGT_MTD_BAL")
						PBTOperationServiceAndInvestmentCol3=rRs("ACT_MTD_BAL_PY1")
						PBTOperationServiceAndInvestmentCol4=rRs("ACT_YTD_BAL")
						PBTOperationServiceAndInvestmentCol5=rRs("BGT_YTD_BAL")
						PBTOperationServiceAndInvestmentCol6=rRs("ACT_YTD_BAL_PY1")
						PBTOperationServiceAndInvestmentCol7=rRs("ACT_YTD_BAL_PY2")
						PBTOperationServiceAndInvestmentCol8=rRs("ACT_YTD_BAL_PY3")
						PBTOperationServiceAndInvestmentCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
				End if
	
				If rRs("DESCCD") ="M5" Then
					If rRs("Sector") = "Manufacturing" Then
						M5ManufacturingCol1=rRs("ACT_MTD_BAL")
						M5ManufacturingCol2=rRs("BGT_MTD_BAL")
						M5ManufacturingCol3=rRs("ACT_MTD_BAL_PY1")
						M5ManufacturingCol4=rRs("ACT_YTD_BAL")
						M5ManufacturingCol5=rRs("BGT_YTD_BAL")
						M5ManufacturingCol6=rRs("ACT_YTD_BAL_PY1")
						M5ManufacturingCol7=rRs("ACT_YTD_BAL_PY2")
						M5ManufacturingCol8=rRs("ACT_YTD_BAL_PY3")
						M5ManufacturingCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
					If rRs("Sector") = "Trading And Contracting" Then
						M5TradingAndContractingCol1=rRs("ACT_MTD_BAL")
						M5TradingAndContractingCol2=rRs("BGT_MTD_BAL")
						M5TradingAndContractingCol3=rRs("ACT_MTD_BAL_PY1")
						M5TradingAndContractingCol4=rRs("ACT_YTD_BAL")
						M5TradingAndContractingCol5=rRs("BGT_YTD_BAL")
						M5TradingAndContractingCol6=rRs("ACT_YTD_BAL_PY1")
						M5TradingAndContractingCol7=rRs("ACT_YTD_BAL_PY2")
						M5TradingAndContractingCol8=rRs("ACT_YTD_BAL_PY3")
						M5TradingAndContractingCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
					If rRs("Sector") = "Service And Investment" Then
						M5ServiceAndInvestmentCol1=rRs("ACT_MTD_BAL")
						M5ServiceAndInvestmentCol2=rRs("BGT_MTD_BAL")
						M5ServiceAndInvestmentCol3=rRs("ACT_MTD_BAL_PY1")
						M5ServiceAndInvestmentCol4=rRs("ACT_YTD_BAL")
						M5ServiceAndInvestmentCol5=rRs("BGT_YTD_BAL")
						M5ServiceAndInvestmentCol6=rRs("ACT_YTD_BAL_PY1")
						M5ServiceAndInvestmentCol7=rRs("ACT_YTD_BAL_PY2")
						M5ServiceAndInvestmentCol8=rRs("ACT_YTD_BAL_PY3")
						M5ServiceAndInvestmentCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
				End if
				If rRs("DESCCD") ="E5" Then
					If rRs("Sector") = "Manufacturing" Then
						E5ManufacturingCol1=rRs("ACT_MTD_BAL")
						E5ManufacturingCol2=rRs("BGT_MTD_BAL")
						E5ManufacturingCol3=rRs("ACT_MTD_BAL_PY1")
						E5ManufacturingCol4=rRs("ACT_YTD_BAL")
						E5ManufacturingCol5=rRs("BGT_YTD_BAL")
						E5ManufacturingCol6=rRs("ACT_YTD_BAL_PY1")
						E5ManufacturingCol7=rRs("ACT_YTD_BAL_PY2")
						E5ManufacturingCol8=rRs("ACT_YTD_BAL_PY3")
						E5ManufacturingCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
					If rRs("Sector") = "Trading And Contracting" Then
						E5TradingAndContractingCol1=rRs("ACT_MTD_BAL")
						E5TradingAndContractingCol2=rRs("BGT_MTD_BAL")
						E5TradingAndContractingCol3=rRs("ACT_MTD_BAL_PY1")
						E5TradingAndContractingCol4=rRs("ACT_YTD_BAL")
						E5TradingAndContractingCol5=rRs("BGT_YTD_BAL")
						E5TradingAndContractingCol6=rRs("ACT_YTD_BAL_PY1")
						E5TradingAndContractingCol7=rRs("ACT_YTD_BAL_PY2")
						E5TradingAndContractingCol8=rRs("ACT_YTD_BAL_PY3")
						E5TradingAndContractingCol9=rRs("ACT_YTD_BAL_PY4")
					End If 
					If rRs("Sector") = "Service And Investment" Then
						E5ServiceAndInvestmentCol1=rRs("ACT_MTD_BAL")
						E5ServiceAndInvestmentCol2=rRs("BGT_MTD_BAL")
						E5ServiceAndInvestmentCol3=rRs("ACT_MTD_BAL_PY1")
						E5ServiceAndInvestmentCol4=rRs("ACT_YTD_BAL")
						E5ServiceAndInvestmentCol5=rRs("BGT_YTD_BAL")
						E5ServiceAndInvestmentCol6=rRs("ACT_YTD_BAL_PY1")
						E5ServiceAndInvestmentCol7=rRs("ACT_YTD_BAL_PY2")
						E5ServiceAndInvestmentCol8=rRs("ACT_YTD_BAL_PY3")
						E5ServiceAndInvestmentCol9=rRs("ACT_YTD_BAL_PY4")
					End If 				
				End if
				rRs.MoveNext
				
			Loop
			
			'
			
'response.write "<br>" &	SalesCol2 
'response.write "<br>" & CommisionIncomeCol2 
'response.write "<br>" & OtherIncome_LossCol2

			' "Revenue" printing start
			sSql = sSql & "<Tr >"
			sSql = sSql & "<TD class='RepfontStylePage17SpecialColor' style='padding: 2px 2px 2px 2px;' class='no-top-border' >"
			sSql = sSql & "Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-border_thick' colspan=3>"
			sSql = sSql & "&nbsp"
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border' >"
			sSql = sSql & "Manufacturing "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol1-RevenueManufacturingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol1-RevenueManufacturingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol4-RevenueManufacturingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol4-RevenueManufacturingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol7,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol8,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-border' align=right>"
			sSql = sSql & formatnumber(RevenueManufacturingCol9,0,,(-1)) 
			sSql = sSql & "</TD>"


			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' >"
			sSql = sSql & "Trading & Contracting "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol1-RevenueTradingAndContractingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol1-RevenueTradingAndContractingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol4-RevenueTradingAndContractingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol4-RevenueTradingAndContractingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick'align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol7,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol8,0,,(-1)) 
			sSql = sSql & "</TD>"
			
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-border'align=right>"
			sSql = sSql & formatnumber(RevenueTradingAndContractingCol9,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'>"
			sSql = sSql & "Service & Investment"
			sSql = sSql & "</TD class='no-top-right-border'>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border' align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol1-RevenueServiceAndInvestmentCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol1-RevenueServiceAndInvestmentCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol4-RevenueServiceAndInvestmentCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol4-RevenueServiceAndInvestmentCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol7,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol8,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(RevenueServiceAndInvestmentCol9,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepfontStyleCalculatedFiledPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border' ><b>"
			sSql = sSql & "Total "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol1 + RevenueTradingAndContractingCol1 + RevenueServiceAndInvestmentCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol2 + RevenueTradingAndContractingCol2 + RevenueServiceAndInvestmentCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol3 + RevenueTradingAndContractingCol3 + RevenueServiceAndInvestmentCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((RevenueManufacturingCol1-RevenueManufacturingCol2) + (RevenueTradingAndContractingCol1-RevenueTradingAndContractingCol2) +  (RevenueServiceAndInvestmentCol1-RevenueServiceAndInvestmentCol2) ,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((RevenueManufacturingCol1-RevenueManufacturingCol3) + (RevenueTradingAndContractingCol1-RevenueTradingAndContractingCol3) +  (RevenueServiceAndInvestmentCol1-RevenueServiceAndInvestmentCol3) ,0,,(-1)) 
			sSql = sSql & "</TD>"


			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol4 + RevenueTradingAndContractingCol4 + RevenueServiceAndInvestmentCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol5 + RevenueTradingAndContractingCol5 + RevenueServiceAndInvestmentCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol6 + RevenueTradingAndContractingCol6 + RevenueServiceAndInvestmentCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((RevenueManufacturingCol4-RevenueManufacturingCol5) + (RevenueTradingAndContractingCol4-RevenueTradingAndContractingCol5) +  (RevenueServiceAndInvestmentCol4-RevenueServiceAndInvestmentCol5) ,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((RevenueManufacturingCol4-RevenueManufacturingCol6) + (RevenueTradingAndContractingCol4-RevenueTradingAndContractingCol6) +  (RevenueServiceAndInvestmentCol4-RevenueServiceAndInvestmentCol6) ,0,,(-1)) 
			sSql = sSql & "</TD>"




			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol7 + RevenueTradingAndContractingCol7 + RevenueServiceAndInvestmentCol7,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol8 + RevenueTradingAndContractingCol8 + RevenueServiceAndInvestmentCol8,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-border'  align=right><b>"
			sSql = sSql & formatnumber(RevenueManufacturingCol9 + RevenueTradingAndContractingCol9 + RevenueServiceAndInvestmentCol9,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			
			' "Revenue" printing End


			' "PBTOperation" printing start
			sSql = sSql & "<Tr >"
			sSql = sSql & "<TD Class=RepfontStylePage17SpecialColor style='padding: 2px 2px 2px 2px;'  class='no-top-border'  >"
			sSql = sSql & "PBT-Operations"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-border_thick' colspan=3>"

			sSql = sSql & "&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  >"
			sSql = sSql & "Manufacturing "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol1-PBTOperationManufacturingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol1-PBTOperationManufacturingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol4-PBTOperationManufacturingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol4-PBTOperationManufacturingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol7,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol8,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol9,0,,(-1)) 
			sSql = sSql & "</TD>"


			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  >"
			sSql = sSql & "Trading & Contracting "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol1-PBTOperationTradingAndContractingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol1-PBTOperationTradingAndContractingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol4-PBTOperationTradingAndContractingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol4-PBTOperationTradingAndContractingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol7,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol8,0,,(-1)) 
			sSql = sSql & "</TD>"
			
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationTradingAndContractingCol9,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  >"
			sSql = sSql & "Service & Investment "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol1-PBTOperationServiceAndInvestmentCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol1-PBTOperationServiceAndInvestmentCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol4-PBTOperationServiceAndInvestmentCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol4-PBTOperationServiceAndInvestmentCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol7,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol8,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(PBTOperationServiceAndInvestmentCol9,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepfontStyleCalculatedFiledPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  ><b>"
			sSql = sSql & "Total "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol1 + PBTOperationTradingAndContractingCol1 + PBTOperationServiceAndInvestmentCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol2 + PBTOperationTradingAndContractingCol2 + PBTOperationServiceAndInvestmentCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol3 + PBTOperationTradingAndContractingCol3 + PBTOperationServiceAndInvestmentCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol1-PBTOperationManufacturingCol2) + (PBTOperationTradingAndContractingCol1-PBTOperationTradingAndContractingCol2) +  (PBTOperationServiceAndInvestmentCol1-PBTOperationServiceAndInvestmentCol2) ,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol1-PBTOperationManufacturingCol3) + (PBTOperationTradingAndContractingCol1-PBTOperationTradingAndContractingCol3) +  (PBTOperationServiceAndInvestmentCol1-PBTOperationServiceAndInvestmentCol3) ,0,,(-1)) 
			sSql = sSql & "</TD>"


			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol4 + PBTOperationTradingAndContractingCol4 + PBTOperationServiceAndInvestmentCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol5 + PBTOperationTradingAndContractingCol5 + PBTOperationServiceAndInvestmentCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol6 + PBTOperationTradingAndContractingCol6 + PBTOperationServiceAndInvestmentCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol4-PBTOperationManufacturingCol5) + (PBTOperationTradingAndContractingCol4-PBTOperationTradingAndContractingCol5) +  (PBTOperationServiceAndInvestmentCol4-PBTOperationServiceAndInvestmentCol5) ,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol4-PBTOperationManufacturingCol6) + (PBTOperationTradingAndContractingCol4-PBTOperationTradingAndContractingCol6) +  (PBTOperationServiceAndInvestmentCol4-PBTOperationServiceAndInvestmentCol6) ,0,,(-1)) 
			sSql = sSql & "</TD>"




			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol7 + PBTOperationTradingAndContractingCol7 + PBTOperationServiceAndInvestmentCol7,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol8 + PBTOperationTradingAndContractingCol8 + PBTOperationServiceAndInvestmentCol8,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right><b>"
			sSql = sSql & formatnumber(PBTOperationManufacturingCol9 + PBTOperationTradingAndContractingCol9 + PBTOperationServiceAndInvestmentCol9,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			
			' "PBTOperation" printing End



			' "CashPBT-Operation" printing start

			sSql = sSql & "<Tr >"
			sSql = sSql & "<TD Class=RepfontStylePage17SpecialColor style='padding: 2px 2px 2px 2px;'  class='no-top-border'  >"
			sSql = sSql & "Cash PBT-Operations"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-border_thick' colspan=3>"

			sSql = sSql & "&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"

			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  >"
			sSql = sSql & "Manufacturing "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol1-E5ManufacturingCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol2-E5ManufacturingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol3-E5ManufacturingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5ManufacturingCol1-E5ManufacturingCol1)-(M5ManufacturingCol2-E5ManufacturingCol2),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5ManufacturingCol1 -E5ManufacturingCol1)-(M5ManufacturingCol3-E5ManufacturingCol3),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol4-E5ManufacturingCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol5-E5ManufacturingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol6-E5ManufacturingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5ManufacturingCol4-E5ManufacturingCol4)- (M5ManufacturingCol5-E5ManufacturingCol5),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5ManufacturingCol4-E5ManufacturingCol4)- (M5ManufacturingCol6-E5ManufacturingCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol7 - E5ManufacturingCol7,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol8 - E5ManufacturingCol8 ,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(M5ManufacturingCol9 - E5ManufacturingCol9 ,0,,(-1)) 
			sSql = sSql & "</TD>"


			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  >"
			sSql = sSql & "Trading & Contracting "
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol1-E5TradingAndContractingCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol2-E5TradingAndContractingCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol3-E5TradingAndContractingCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5TradingAndContractingCol1-E5TradingAndContractingCol1)-(M5TradingAndContractingCol2-E5TradingAndContractingCol2),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5TradingAndContractingCol1 -E5TradingAndContractingCol1)-(M5TradingAndContractingCol3-E5TradingAndContractingCol3),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol4-E5TradingAndContractingCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol5-E5TradingAndContractingCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol6-E5TradingAndContractingCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5TradingAndContractingCol4-E5TradingAndContractingCol4)- (M5TradingAndContractingCol5-E5TradingAndContractingCol5),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5TradingAndContractingCol4-E5TradingAndContractingCol4)- (M5TradingAndContractingCol6-E5TradingAndContractingCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol7 - E5TradingAndContractingCol7,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol8 - E5TradingAndContractingCol8 ,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(M5TradingAndContractingCol9 - E5TradingAndContractingCol9 ,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  >"
			sSql = sSql & "Service & Investment "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol1-E5ServiceAndInvestmentCol1,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol2-E5ServiceAndInvestmentCol2,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol3-E5ServiceAndInvestmentCol3,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5ServiceAndInvestmentCol1-E5ServiceAndInvestmentCol1)-(M5ServiceAndInvestmentCol2-E5ServiceAndInvestmentCol2),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5ServiceAndInvestmentCol1 -E5ServiceAndInvestmentCol1)-(M5ServiceAndInvestmentCol3-E5ServiceAndInvestmentCol3),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol4-E5ServiceAndInvestmentCol4,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol5-E5ServiceAndInvestmentCol5,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol6-E5ServiceAndInvestmentCol6,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5ServiceAndInvestmentCol4-E5ServiceAndInvestmentCol4)- (M5ServiceAndInvestmentCol5-E5ServiceAndInvestmentCol5),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((M5ServiceAndInvestmentCol4-E5ServiceAndInvestmentCol4)- (M5ServiceAndInvestmentCol6-E5ServiceAndInvestmentCol6),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol7 - E5ServiceAndInvestmentCol7,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol8 - E5ServiceAndInvestmentCol8 ,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(M5ServiceAndInvestmentCol9 - E5ServiceAndInvestmentCol9 ,0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepfontStyleCalculatedFiledPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  ><b>"
			sSql = sSql & "Total "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol1-E5ManufacturingCol1) + (M5TradingAndContractingCol1 - E5TradingAndContractingCol1) + (M5ServiceAndInvestmentCol1 - E5ServiceAndInvestmentCol1),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol2-E5ManufacturingCol2) + (M5TradingAndContractingCol2 - E5TradingAndContractingCol2) + (M5ServiceAndInvestmentCol2 - E5ServiceAndInvestmentCol2),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol3-E5ManufacturingCol3) + (M5TradingAndContractingCol3 - E5TradingAndContractingCol3) + (M5ServiceAndInvestmentCol3 - E5ServiceAndInvestmentCol3),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(  ((M5ManufacturingCol1-E5ManufacturingCol1)-(M5ManufacturingCol2-E5ManufacturingCol2)) + ((M5TradingAndContractingCol1-E5TradingAndContractingCol1) - (M5TradingAndContractingCol2-E5TradingAndContractingCol2)) +  ((M5ServiceAndInvestmentCol1-E5ServiceAndInvestmentCol1)-(M5ServiceAndInvestmentCol2-E5ServiceAndInvestmentCol2)) ,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(  ((M5ManufacturingCol1-E5ManufacturingCol1)-(M5ManufacturingCol3-E5ManufacturingCol3)) + ((M5TradingAndContractingCol1-E5TradingAndContractingCol1) - (M5TradingAndContractingCol3-E5TradingAndContractingCol3)) +  ((M5ServiceAndInvestmentCol1-E5ServiceAndInvestmentCol1)-(M5ServiceAndInvestmentCol3-E5ServiceAndInvestmentCol3)) ,0,,(-1)) 
			sSql = sSql & "</TD>"


			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol4-E5ManufacturingCol4) + (M5TradingAndContractingCol4 - E5TradingAndContractingCol4) + (M5ServiceAndInvestmentCol4 - E5ServiceAndInvestmentCol4),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol5-E5ManufacturingCol5) + (M5TradingAndContractingCol5 - E5TradingAndContractingCol5) + (M5ServiceAndInvestmentCol5 - E5ServiceAndInvestmentCol5),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol6-E5ManufacturingCol6) + (M5TradingAndContractingCol6 - E5TradingAndContractingCol6) + (M5ServiceAndInvestmentCol6 - E5ServiceAndInvestmentCol6),0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(  ((M5ManufacturingCol4-E5ManufacturingCol4)-(M5ManufacturingCol5-E5ManufacturingCol5)) + ((M5TradingAndContractingCol4-E5TradingAndContractingCol4) - (M5TradingAndContractingCol5-E5TradingAndContractingCol5)) +  ((M5ServiceAndInvestmentCol4-E5ServiceAndInvestmentCol4)-(M5ServiceAndInvestmentCol5-E5ServiceAndInvestmentCol5)) ,0,,(-1)) 
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(  ((M5ManufacturingCol4-E5ManufacturingCol4)-(M5ManufacturingCol6-E5ManufacturingCol6)) + ((M5TradingAndContractingCol4-E5TradingAndContractingCol4) - (M5TradingAndContractingCol6-E5TradingAndContractingCol6)) +  ((M5ServiceAndInvestmentCol4-E5ServiceAndInvestmentCol4)-(M5ServiceAndInvestmentCol6-E5ServiceAndInvestmentCol6)) ,0,,(-1)) 
			sSql = sSql & "</TD>"




			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol7-E5ManufacturingCol7) + (M5TradingAndContractingCol7 - E5TradingAndContractingCol7) + (M5ServiceAndInvestmentCol7 - E5ServiceAndInvestmentCol7),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol8-E5ManufacturingCol8) + (M5TradingAndContractingCol8 - E5TradingAndContractingCol8) + (M5ServiceAndInvestmentCol8 - E5ServiceAndInvestmentCol8),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right><b>"
			sSql = sSql & FormatNumber((M5ManufacturingCol9-E5ManufacturingCol9) + (M5TradingAndContractingCol9 - E5TradingAndContractingCol9) + (M5ServiceAndInvestmentCol9 - E5ServiceAndInvestmentCol9),0,,(-1)) 
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			
			' "M5" printing End



			' "PBTOperation % Revenue" printing start

			sSql = sSql & "<Tr >"
			sSql = sSql & "<TD Class=RepfontStylePage17SpecialColor style='padding: 2px 2px 2px 2px;'  class='no-top-border'  >"
			sSql = sSql & "PBT-Operations % to Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-border_thick' colspan=3>"

			sSql = sSql & "&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"


			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  >"
			sSql = sSql & "Manufacturing "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			'response.write " RevenueManufacturingCol1 = " & RevenueManufacturingCol1 & "<br>"
			If RevenueManufacturingCol1 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol1/RevenueManufacturingCol1)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If
			'response.write " RevenueManufacturingCol2 = " & RevenueManufacturingCol2 & "<br>"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"

			If RevenueManufacturingCol2 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol2/RevenueManufacturingCol2)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If

			
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			If RevenueManufacturingCol3 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol3/RevenueManufacturingCol3)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If

			
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD  style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD  style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"

			If RevenueManufacturingCol4 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol4/RevenueManufacturingCol4)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If

			
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			If RevenueManufacturingCol5 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol5/RevenueManufacturingCol5)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If

			
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			If RevenueManufacturingCol6 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol6/RevenueManufacturingCol6)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If

			
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD  style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD  style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			If RevenueManufacturingCol7 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol7/RevenueManufacturingCol7)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If

			
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			If RevenueManufacturingCol8 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol8/RevenueManufacturingCol8)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If

			
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-border'  align=right>"
			If RevenueManufacturingCol9 <> 0 then
				sSql = sSql & formatnumber((PBTOperationManufacturingCol9/RevenueManufacturingCol9)*100,1,,(-1)) & "%"
			Else
				sSql = sSql & "&nbsp;"
			End If

			
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  >"
			sSql = sSql & "Trading & Contracting"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol1/RevenueTradingAndContractingCol1)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol2/RevenueTradingAndContractingCol2)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol3/RevenueTradingAndContractingCol3)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD  style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD  style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol4/RevenueTradingAndContractingCol4)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol5/RevenueTradingAndContractingCol5)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'   class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol6/RevenueTradingAndContractingCol6)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol7/RevenueTradingAndContractingCol7)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol8/RevenueTradingAndContractingCol8)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationTradingAndContractingCol9/RevenueTradingAndContractingCol9)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  >"
			sSql = sSql & "Service & Investment"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol1/RevenueServiceAndInvestmentCol1)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol2/RevenueServiceAndInvestmentCol2)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol3/RevenueServiceAndInvestmentCol3)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol4/RevenueServiceAndInvestmentCol4)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol5/RevenueServiceAndInvestmentCol5)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol6/RevenueServiceAndInvestmentCol6)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol7/RevenueServiceAndInvestmentCol7)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol8/RevenueServiceAndInvestmentCol8)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-border'  align=right>"
			sSql = sSql & formatnumber((PBTOperationServiceAndInvestmentCol9/RevenueServiceAndInvestmentCol9)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			


			'-*-
			sSql = sSql & "<Tr Class=RepfontStyleCalculatedFiledPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  ><b>"
			sSql = sSql & "Total "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol1 + PBTOperationTradingAndContractingCol1 + PBTOperationServiceAndInvestmentCol1)/(RevenueManufacturingCol1 + RevenueTradingAndContractingCol1 + RevenueServiceAndInvestmentCol1)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol2 + PBTOperationTradingAndContractingCol2 + PBTOperationServiceAndInvestmentCol2)/(RevenueManufacturingCol2 + RevenueTradingAndContractingCol2 + RevenueServiceAndInvestmentCol2)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol3 + PBTOperationTradingAndContractingCol3 + PBTOperationServiceAndInvestmentCol3)/(RevenueManufacturingCol3 + RevenueTradingAndContractingCol3 + RevenueServiceAndInvestmentCol3)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"


			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol4 + PBTOperationTradingAndContractingCol4 + PBTOperationServiceAndInvestmentCol4)/(RevenueManufacturingCol4 + RevenueTradingAndContractingCol4 + RevenueServiceAndInvestmentCol4)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol5 + PBTOperationTradingAndContractingCol5 + PBTOperationServiceAndInvestmentCol5)/(RevenueManufacturingCol5 + RevenueTradingAndContractingCol5 + RevenueServiceAndInvestmentCol5)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol6 + PBTOperationTradingAndContractingCol6 + PBTOperationServiceAndInvestmentCol6)/(RevenueManufacturingCol6 + RevenueTradingAndContractingCol6 + RevenueServiceAndInvestmentCol6)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol7 + PBTOperationTradingAndContractingCol7 + PBTOperationServiceAndInvestmentCol7)/(RevenueManufacturingCol7 + RevenueTradingAndContractingCol7 + RevenueServiceAndInvestmentCol7)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol8 + PBTOperationTradingAndContractingCol8 + PBTOperationServiceAndInvestmentCol8)/(RevenueManufacturingCol8 + RevenueTradingAndContractingCol8 + RevenueServiceAndInvestmentCol8)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-border'  align=right><b>"
			sSql = sSql & formatnumber((PBTOperationManufacturingCol9 + PBTOperationTradingAndContractingCol9 + PBTOperationServiceAndInvestmentCol9)/(RevenueManufacturingCol9 + RevenueTradingAndContractingCol9 + RevenueServiceAndInvestmentCol9)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			' "PBTOperation % Revenue" printing End






			' "CashPBTOperation % Revenue" printing start
			sSql = sSql & "<Tr >"
			sSql = sSql & "<TD Class=RepfontStylePage17SpecialColor style='padding: 2px 2px 2px 2px;'  class='no-top-border'  >"
			sSql = sSql & "Cash PBT-Operations % to Revenue"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border_thick' colspan=5>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-border_thick' colspan=3>"

			sSql = sSql & "&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"

			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  >"
			sSql = sSql & "Manufacturing "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol1-E5ManufacturingCol1)/RevenueManufacturingCol1)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol2-E5ManufacturingCol2)/RevenueManufacturingCol2)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol3-E5ManufacturingCol3)/RevenueManufacturingCol3)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol4-E5ManufacturingCol4)/RevenueManufacturingCol4)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol5-E5ManufacturingCol5)/RevenueManufacturingCol5)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol6-E5ManufacturingCol6)/RevenueManufacturingCol6)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol7-E5ManufacturingCol7)/RevenueManufacturingCol7)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol8-E5ManufacturingCol8)/RevenueManufacturingCol8)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(((M5ManufacturingCol9-E5ManufacturingCol9)/RevenueManufacturingCol9)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  >"
			sSql = sSql & "Trading & Contracting"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol1-E5TradingAndContractingCol1)/RevenueTradingAndContractingCol1)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol2-E5TradingAndContractingCol2)/RevenueTradingAndContractingCol2)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol3-E5TradingAndContractingCol3)/RevenueTradingAndContractingCol3)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol4-E5TradingAndContractingCol4)/RevenueTradingAndContractingCol4)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol5-E5TradingAndContractingCol5)/RevenueTradingAndContractingCol5)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol6-E5TradingAndContractingCol6)/RevenueTradingAndContractingCol6)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol7-E5TradingAndContractingCol7)/RevenueTradingAndContractingCol7)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol8-E5TradingAndContractingCol8)/RevenueTradingAndContractingCol8)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(((M5TradingAndContractingCol9-E5TradingAndContractingCol9)/RevenueTradingAndContractingCol9)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"
			'-*-
			sSql = sSql & "<Tr Class=RepContentFontPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  >"
			sSql = sSql & "Service & Investment"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol1-E5ServiceAndInvestmentCol1)/RevenueServiceAndInvestmentCol1)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol2-E5ServiceAndInvestmentCol2)/RevenueServiceAndInvestmentCol2)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol3-E5ServiceAndInvestmentCol3)/RevenueServiceAndInvestmentCol3)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol4-E5ServiceAndInvestmentCol4)/RevenueServiceAndInvestmentCol4)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol5-E5ServiceAndInvestmentCol5)/RevenueServiceAndInvestmentCol5)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol6-E5ServiceAndInvestmentCol6)/RevenueServiceAndInvestmentCol6)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol7-E5ServiceAndInvestmentCol7)/RevenueServiceAndInvestmentCol7)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol8-E5ServiceAndInvestmentCol8)/RevenueServiceAndInvestmentCol8)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-border'  align=right>"
			sSql = sSql & formatnumber(((M5ServiceAndInvestmentCol9-E5ServiceAndInvestmentCol9)/RevenueServiceAndInvestmentCol9)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			


			'-*-
			sSql = sSql & "<Tr Class=RepfontStyleCalculatedFiledPage17>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  ><b>"
			sSql = sSql & "Total "
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol1-E5ManufacturingCol1) + (M5TradingAndContractingCol1-E5TradingAndContractingCol1) + (M5ServiceAndInvestmentCol1-E5ServiceAndInvestmentCol1))/(RevenueManufacturingCol1 + RevenueTradingAndContractingCol1 + RevenueServiceAndInvestmentCol1)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol2-E5ManufacturingCol2) + (M5TradingAndContractingCol2-E5TradingAndContractingCol2) + (M5ServiceAndInvestmentCol2-E5ServiceAndInvestmentCol2))/(RevenueManufacturingCol2 + RevenueTradingAndContractingCol2 + RevenueServiceAndInvestmentCol2)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol3-E5ManufacturingCol3) + (M5TradingAndContractingCol3-E5TradingAndContractingCol3) + (M5ServiceAndInvestmentCol3-E5ServiceAndInvestmentCol3))/(RevenueManufacturingCol3 + RevenueTradingAndContractingCol3 + RevenueServiceAndInvestmentCol3)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"


			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol4-E5ManufacturingCol4) + (M5TradingAndContractingCol4-E5TradingAndContractingCol4) + (M5ServiceAndInvestmentCol4-E5ServiceAndInvestmentCol4))/(RevenueManufacturingCol4 + RevenueTradingAndContractingCol4 + RevenueServiceAndInvestmentCol4)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol5-E5ManufacturingCol5) + (M5TradingAndContractingCol5-E5TradingAndContractingCol5) + (M5ServiceAndInvestmentCol5-E5ServiceAndInvestmentCol5))/(RevenueManufacturingCol5 + RevenueTradingAndContractingCol5 + RevenueServiceAndInvestmentCol5)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol6-E5ManufacturingCol6) + (M5TradingAndContractingCol6-E5TradingAndContractingCol6) + (M5ServiceAndInvestmentCol6-E5ServiceAndInvestmentCol6))/(RevenueManufacturingCol6 + RevenueTradingAndContractingCol6 + RevenueServiceAndInvestmentCol6)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' class='no-top-right-border'  align=right>&nbsp;"
			sSql = sSql & "</TD>"

			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border_thick'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol7-E5ManufacturingCol7) + (M5TradingAndContractingCol7-E5TradingAndContractingCol7) + (M5ServiceAndInvestmentCol7-E5ServiceAndInvestmentCol7))/(RevenueManufacturingCol7 + RevenueTradingAndContractingCol7 + RevenueServiceAndInvestmentCol7)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-right-border'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol8-E5ManufacturingCol8) + (M5TradingAndContractingCol8-E5TradingAndContractingCol8) + (M5ServiceAndInvestmentCol8-E5ServiceAndInvestmentCol8))/(RevenueManufacturingCol8 + RevenueTradingAndContractingCol8 + RevenueServiceAndInvestmentCol8)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;'  class='no-top-border'  align=right><b>"
			sSql = sSql & formatnumber(((M5ManufacturingCol9-E5ManufacturingCol9) + (M5TradingAndContractingCol9-E5TradingAndContractingCol9) + (M5ServiceAndInvestmentCol9-E5ServiceAndInvestmentCol9))/(RevenueManufacturingCol9 + RevenueTradingAndContractingCol9 + RevenueServiceAndInvestmentCol9)*100,1,,(-1)) & "%"
			sSql = sSql & "</TD>"
			sSql = sSql & "</Tr>"

			' "CashPBTOperation % Revenue" printing End

			sSql = sSql & "<TR Class=RepfontStyleReportRemark>"
			sSql = sSql & "<TD style='padding: 2px 2px 2px 2px;' colspan=14>*Revenue includes Sales, Commission and Other Income."
			sSql = sSql & "</TD>"
			sSql = sSql & "</TR>"	
			
			sSql = sSql & "</Table>"

			'		
			Response.write sSql 

		Else
			Response.write "<Tr Class=cellStyle1><Td Align=Center Colspan=27><b>**********No Record Fetched (A)**********</Td>"
		End if		
		rRs.close
	
	End function

	Private Function page15(period)


		'

		sSql = "SELECT mistxn.*, COMPANY.CCOMPNAME,COMPANY.CATEGORY,COMPANY.ccompnamesn2 "
		sSql = sSql & " FROM (COMPANY RIGHT JOIN mistxn ON COMPANY.CCOMPCODE = mistxn.ccompcode) "
		sSql = sSql & " where  Trim(MisTxn.cMonth)='" & TheReportMonth & "' AND Trim(MisTxn.cYear)='" & TheReportYear & "' And CATEGORY <>'Others' ORDER BY  COMPANY.cCompname , mistxn.DescCD"
							

		'Response.write sSql & "<BR>"
		'response.end


		rRs.Open Trim(sSql),cCn
		Content=""
		SRLNO = 0
		PageNumber=1
		Line_Count=1
		CompID=""
		CompanyName=""

		Profit_Before_Tax_ACT_MTD = 0
		Add_Commission_Income_ACT_MTD = 0
		Add_Investment_Income_ACT_MTD = 0

		
		If Not rRs.Eof Then
			'

			sSql = ""
			cMM=Trim(rRs("cMonth"))
			cYY=Trim(rRs("cYear"))

			Profit_Before_Tax_ACT_MTD = 0
			Add_Commission_Income_ACT_MTD = 0
			Add_Investment_Income_ACT_MTD = 0


			Total_Sales=0
			Total_Costofgoodssold=0
			Total_GrossMargin=0
			Total_CommissionIncome=0
			Total_InvestmentIncome=0
			Total_OtherIncome=0
			Total_GrossIncome=0
			Total_Overheads=0
			Total_OperatingIncome=0
			Total_LessFinanceCost=0
			Total_ProvisionForInvestment_Repeat=0
			Total_ProfitBeforeTax=0
			Total_AddDepreciation=0
			Total_CashProfitLoss=0
			Total_TotalRevenue=0
			Total_PBTFromOperation =0
			Total_DividendToShareholders=0
			Total_PBTFromOperation_historicdata=0
			'
			Total_BudgetSales=0
			Total_BudgetPBTFromOperation=0
			Total_VariancePBTFromOperation=0
			'

			Add_Net_Investement_Income_col=0
			Cash_Profit_Before_Tax_col=0
			Total_CashPBTFromOperation=0


			sSql = sSql & "<Table border=0  width='100%' cellPadding=0 cellspacing=0>"
		
			sSql = sSql & "<Tr>"			
			If (period = "M") then
				sSql = sSql & "<Td align='Left' class='GreenTransparentHeadingNormal' >Company-wise Performance Summary of Group Results for the period:&nbsp;" & FullMonthsArr(cint(TheReportMonth)-1) & " " & TheReportYear & "&nbsp;(MTD)"
			Else
				sSql = sSql & "<Td align='left' class='GreenTransparentHeadingNormal' >Company-wise Performance Summary of Group Results for the period:&nbsp;" & FullMonthsArr(cint(TheReportMonth)-1) & " " & TheReportYear & "&nbsp;(YTD)"
			End if 
			sSql = sSql & "</Td>"
			sSql = sSql & "	<Td Align=Right  class='RepfontStyleCommonTransparentHeadingSmall'  > Amount (RO '000)</Td>"
			sSql = sSql & "</Tr>"

			sSql = sSql & "<Tr >"
			sSql = sSql & "<Td Align=center valign=top Colspan=2><Table Border=0 Cellspacing=0 Cellpadding=0 Width='100%'>"
			sSql = sSql & "<tr class='SubRepHeaderPage15And16'><Td Align='center' class='no-right-border'><b>OGC</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Sales</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Budget<br>Sales</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>COGS</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Gross<br>Margin</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Comm.<br>Inc</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Invst.<br>Inc</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Oth<br>Inc/<br>(Loss)</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Gross<br>Profit</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Less<br>OH</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Opr<br>Inc</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Less<br>Fin Cost<br>/(Inc)</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Less<br>Prov For<br> Invest</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>PBT</td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Add<br>Depre</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Cash<br>PBT</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Gross<br>Margin<br>%</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>PBT<br>%</td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>PBT<br>From<br>Ops </b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Cash<br>PBT<br>From<br>Ops </b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Divi<br>Paid<br>Out</b></td>"
			sSql = sSql & "<Td Align='center' class='no-right-border' valign='center'><b>Budget<BR> PBT. Ops </b></td>"
			sSql = sSql & "<Td Align='center' class='border-all' valign='center'><b>Variance<br>PBT. Ops </b></td>"

			sSql = sSql & "</tr>"
				
			'
			Actual_PBT_Operation=0
			Budget_PBT_Operation=0

			Do while not rRs.EOF
				'
				If (rRs("DESCCD")<>"L7") Then
					SRLNO = SRLNO + 1

					If Line_Count=1 then
					end If
					If (CompID<>Trim(rRs("cCompCode"))) Then
						If CompID<>"" Then
							sSql = sSql & "</Tr>"
							Profit_Before_Tax_ACT_MTD = 0
							Add_Commission_Income_ACT_MTD = 0
							Add_Investment_Income_ACT_MTD = 0
						End If 
						CompID=Trim(rRs("cCompCode"))
						sSql = sSql & "<Tr Class=RepfontStylePage15And16>"
						sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=left nowrap class='no-top-right-border'>" & Trim(rRs("ccompnamesn2")) & "</Td>"
					End If 

				If rRs("desccd") = "E5" Then
					If (period = "M") then
						Add_Net_Investement_Income_col=rRs("ACT_MTD_BAL")
					Else
						Add_Net_Investement_Income_col=rRs("ACT_YTD_BAL")
					End If 
				End if
				
				
				If rRs("desccd") = "M5" Then
					If (period = "M") then
						Cash_Profit_Before_Tax_col=rRs("ACT_MTD_BAL")
					Else
						Cash_Profit_Before_Tax_col=rRs("ACT_YTD_BAL")
					End If 
				End If

					If (rRs("DESCCD")<>"P5" And  rRs("DESCCD")<>"Q5") Then
						If (period = "M") Then
							If (cdbl(rRs("ACT_MTD_BAL")) <> 0) then
								sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right  class='no-top-right-border'>" & formatnumber(rRs("ACT_MTD_BAL"),0,,(-1)) & "</Td>"
							Else
								sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right  class='no-top-right-border'>&nbsp;</Td>"
							End If 
						Else
							If cdbl(rRs("ACT_YTD_BAL")) <> 0 then
								sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right  class='no-top-right-border'>" & formatnumber(rRs("ACT_YTD_BAL"),0,,(-1)) & "</Td>"
							Else
								sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right  class='no-top-right-border'>&nbsp;</Td>"
							End If 
						End if
					End if 
		
					If (rRs("DESCCD")="A5") Then
						If (period = "M") Then
								If cdbl(rRs("BGT_MTD_BAL")) <> 0 then
									sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(rRs("BGT_MTD_BAL"),0,,(-1)) & "</Td>"
								Else
									sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>&nbsp;</Td>"
								End If 
								Total_BudgetSales = Total_BudgetSales + formatnumber(rRs("BGT_MTD_BAL"),0,,(-1))
						Else
								If cdbl(rRs("BGT_YTD_BAL")) <> 0 then
									sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(rRs("BGT_YTD_BAL"),0,,(-1)) & "</Td>"	
								Else
									sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>&nbsp;</Td>"	
								End If 
								Total_BudgetSales = Total_BudgetSales + formatnumber(rRs("BGT_YTD_BAL"),0,,(-1))
						End If
					End if

					
					If (rRs("DESCCD")="R5") Then
						If (period = "M") then
							Budget_PBT_Operation =  rRs("BGT_MTD_BAL")
							Actual_PBT_Operation =  rRs("ACT_MTD_BAL")
							Total_BudgetPBTFromOperation = Total_BudgetPBTFromOperation  + formatnumber(rRs("BGT_MTD_BAL"),0,,(-1))
						Else
							Budget_PBT_Operation =  rRs("BGT_YTD_BAL")
							Actual_PBT_Operation =  rRs("ACT_YTD_BAL")
							Total_BudgetPBTFromOperation = Total_BudgetPBTFromOperation  + formatnumber(rRs("BGT_YTD_BAL"),0,,(-1))
						End If
						If cdbl(Cash_Profit_Before_Tax_col) - CDbl(Add_Net_Investement_Income_col)  <> 0 then
							sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(CDbl(Cash_Profit_Before_Tax_col) -  CDbl(Add_Net_Investement_Income_col),0,,(-1)) & "</Td>"
						Else
							sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>&nbsp;</Td>"
						End if
						Total_CashPBTFromOperation = Total_CashPBTFromOperation + (CDbl(Cash_Profit_Before_Tax_col) -  CDbl(Add_Net_Investement_Income_col))

					End if

					If (rRs("DESCCD")="R6") Then
							If cdbl(Budget_PBT_Operation) <> 0 then
								sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Budget_PBT_Operation,0,,(-1)) & "</Td>"
							Else
								sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>&nbsp;</Td>"
							End if
								'sSql = sSql & "<Td Align=right>" & formatnumber(Actual_PBT_Operation,0,,(-1)) - formatnumber(Budget_PBT_Operation,0,,(-1)) & "</Td>"
							If (cdbl(Actual_PBT_Operation) - cdbl(Budget_PBT_Operation)) <> 0 then
								sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-border'>" & FormatNumber(formatnumber(Actual_PBT_Operation,0,,(-1)) - formatnumber(Budget_PBT_Operation,0,,(-1)) ,0,,(-1)) & "</Td>"
							Else 
								sSql = sSql & "<Td style='padding: 0 2px 0 2px;' Align=right class='no-top-border'>&nbsp;</Td>"
							End if 
								Total_VariancePBTFromOperation = Total_VariancePBTFromOperation + (CDbl(Actual_PBT_Operation) - CDbl(Budget_PBT_Operation))
								'response.write Total_VariancePBTFromOperation & "<br>"
					End if

					If (rRs("DESCCD")="K5") Then
						If (period = "M") then
							Profit_Before_Tax_ACT_MTD = rRs("ACT_MTD_BAL")
						Else 
							Profit_Before_Tax_ACT_MTD = rRs("ACT_YTD_BAL")
						End if
						'response.write "<br>Profit_Before_Tax_ACT_MTD = " & Profit_Before_Tax_ACT_MTD
					End If 

					If (rRs("DESCCD")="D5") Then
						If (period = "M") then
							Add_Commission_Income_ACT_MTD = rRs("ACT_MTD_BAL")
						Else
							Add_Commission_Income_ACT_MTD = rRs("ACT_YTD_BAL")
						End if
						'response.write "<br>Add_Commission_Income_ACT_MTD = " & Add_Commission_Income_ACT_MTD
					End If 
	
					If (rRs("DESCCD")="E5") Then
						If (period = "M") then
							Add_Investment_Income_ACT_MTD = rRs("ACT_MTD_BAL")
						Else
							Add_Investment_Income_ACT_MTD = rRs("ACT_YTD_BAL")
						End if 
						'response.write "<br>Add_Investment_Income_ACT_MTD = " & Add_Investment_Income_ACT_MTD
					End If 
	

					If (rRs("DESCCD")="A5") Then
						If (period = "M") then
							Total_Sales=Total_Sales + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_Sales=Total_Sales + formatnumber(rRs("ACT_YTD_BAL"),0,,(-1)) 
						End if 
					End If 

					If (rRs("DESCCD")="B5") Then
						If (period = "M") then
							Total_Costofgoodssold=Total_Costofgoodssold + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_Costofgoodssold=Total_Costofgoodssold +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 


					If (rRs("DESCCD")="C5") Then
						If (period = "M") then
							Total_GrossMargin=Total_GrossMargin + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_GrossMargin=Total_GrossMargin +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

					If (rRs("DESCCD")="D5") Then
						If (period = "M") then
							Total_CommissionIncome=Total_CommissionIncome + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_CommissionIncome=Total_CommissionIncome +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 


					If (rRs("DESCCD")="E5") Then
						If (period = "M") then
							Total_InvestmentIncome=Total_InvestmentIncome + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_InvestmentIncome=Total_InvestmentIncome +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

					If (rRs("DESCCD")="E6") Then
						If (period = "M") then
							Total_OtherIncome=Total_OtherIncome + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_OtherIncome=Total_OtherIncome +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

					If (rRs("DESCCD")="F5") Then
						If (period = "M") then
							Total_GrossIncome=Total_GrossIncome + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_GrossIncome=Total_GrossIncome +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 
					If (rRs("DESCCD")="G5") Then
						If (period = "M") then
							Total_Overheads=Total_Overheads + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_Overheads=Total_Overheads +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

					If (rRs("DESCCD")="H5") Then
						If (period = "M") then
							Total_OperatingIncome=Total_OperatingIncome + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_OperatingIncome=Total_OperatingIncome +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

					If (rRs("DESCCD")="J5") Then
						If (period = "M") then
							Total_LessFinanceCost=Total_LessFinanceCost + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_LessFinanceCost=Total_LessFinanceCost +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 
					
					If (rRs("DESCCD")="J7") Then
						If (period = "M") then
							Total_ProvisionForInvestment_Repeat=Total_ProvisionForInvestment_Repeat + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_ProvisionForInvestment_Repeat=Total_ProvisionForInvestment_Repeat +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

					If (rRs("DESCCD")="K5") Then
						If (period = "M") then
							Total_ProfitBeforeTax=Total_ProfitBeforeTax + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_ProfitBeforeTax=Total_ProfitBeforeTax +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

					If (rRs("DESCCD")="L5") Then
						If (period = "M") then
							Total_AddDepreciation=Total_AddDepreciation + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_AddDepreciation=Total_AddDepreciation +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 
					
					'Commented on 27Aug2018 Mr. Kamal does not want
					If (rRs("DESCCD")="M5") Then
						If (period = "M") then
							Total_CashProfitLoss=Total_CashProfitLoss + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_CashProfitLoss=Total_CashProfitLoss +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

					
					If (rRs("DESCCD")="Q5") Then
						If (period = "M") then
							Total_TotalRevenue=Total_TotalRevenue + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_TotalRevenue=Total_TotalRevenue +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 


					If (rRs("DESCCD")="R5") Then
						If (period = "M") then
							Total_PBTFromOperation=Total_PBTFromOperation + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_PBTFromOperation=Total_PBTFromOperation +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 
					
					If (rRs("DESCCD")="R6") Then
						If (period = "M") then
							Total_DividendToShareholders=Total_DividendToShareholders + formatnumber(rRs("ACT_MTD_BAL"),0,,(-1))
						Else
							Total_DividendToShareholders=Total_DividendToShareholders +  formatnumber(rRs("ACT_YTD_BAL"),0,,(-1))
						End if 
					End If 

				End if
				rRs.MoveNext
			Loop

' WHY ?
'			if request.querystring("prd") < "201805" then
'				sSql = sSql & "<Td Align=right>" & formatnumber(Profit_Before_Tax_ACT_MTD,cDec) - formatnumber(Add_Investment_Income_ACT_MTD,cDec)  & "</Td></tr>"
'				Total_PBTFromOperation_historicdata = Total_PBTFromOperation_historicdata + formatnumber(Profit_Before_Tax_ACT_MTD,cDec) -  formatnumber(Add_Investment_Income_ACT_MTD,cDec)
'			End if

			sSql = sSql & "<Tr Class=RepfontStyleCalculatedFiledPage15And16>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' class='no-top-right-border'>Total</td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_Sales,0,,(-1))& "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_BudgetSales,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_Costofgoodssold,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_GrossMargin,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_CommissionIncome,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_InvestmentIncome,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_OtherIncome,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_GrossIncome,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_Overheads,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_OperatingIncome,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_LessFinanceCost,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_ProvisionForInvestment_Repeat,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_ProfitBeforeTax,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_AddDepreciation,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_CashProfitLoss,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD class='no-top-right-border'>&nbsp;</td>"
			sSql = sSql & "<TD class='no-top-right-border'>&nbsp;</td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_PBTFromOperation,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_CashPBTFromOperation,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_DividendToShareholders,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-right-border'>" & formatnumber(Total_BudgetPBTFromOperation,0,,(-1)) & "</Td>"
			sSql = sSql & "<TD style='padding: 0 2px 0 2px;' Align=right class='no-top-border'>" & formatnumber(Total_VariancePBTFromOperation,0,,(-1)) & "</Td>"

			sSql = sSql & "</Tr>"
			sSql = sSql & "</table>"
			sSql = sSql & "</Tr>"
			sSql = sSql & "</table>"
			'
			Response.write sSql 
		Else
			Response.write "<Tr Class=cellStyle1><Td Align=Center Colspan=27>No Record Fetched</Td>"
		End if		
		rRs.close

	End function

	Private Function Page14()

		sSql="" 
		sSql= "SELECT  DESCCD, DESCRIPTION, ACT_MTD_BAL , BGT_MTD_BAL,ACT_YTD_BAL , BGT_YTD_BAL,ACT_MTD_BAL_PY1,ACT_YTD_BAL_PY1" 
		sSql= sSql & " FROM DEPTWISEPBT   " 
		sSql= sSql & " where Trim(cMonth)='" & TheReportMonth & "'  " 
		sSql= sSql & " AND Trim(cYear)='" & TheReportYear & "'  " 
		sSql= sSql & " And  DescCd  not in ('C5') "
		sSql= sSql & " order by 1 "

'response.write sSql
'response.end
		rRs.Open Trim(sSql),cCn


		If Not rRs.Eof Then
			'
			A1Desc=""
			A1MC1=0
			A1MC2=0
			A1MC3=0
			A1YC1=0
			A1YC2=0
			A1YC3=0

			A5Desc=""
			A5MC1=0
			A5MC2=0
			A5MC3=0
			A5YC1=0
			A5YC2=0
			A5YC3=0

			B1Desc=""
			B1MC1=0
			B1MC2=0
			B1MC3=0
			B1YC1=0
			B1YC2=0
			B1YC3=0

			B5Desc=""
			B5MC1=0
			B5MC2=0
			B5MC3=0
			B5YC1=0
			B5YC2=0
			B5YC3=0

			C1Desc=""
			C1MC1=0
			C1MC2=0
			C1MC3=0
			C1YC1=0
			C1YC2=0
			C1YC3=0

			C5Desc=""
			C5MC1=0
			C5MC2=0
			C5MC3=0
			C5YC1=0
			C5YC2=0
			C5YC3=0

			E5Desc=""
			E5MC1=0
			E5MC2=0
			E5MC3=0
			E5YC1=0
			E5YC2=0
			E5YC3=0

			K5Desc=""
			K5MC1=0
			K5MC2=0
			K5MC3=0
			K5YC1=0
			K5YC2=0
			K5YC3=0

			R5Desc=""
			R5MC1=0
			R5MC2=0
			R5MC3=0
			R5YC1=0
			R5YC2=0
			R5YC3=0
			Do while not rRs.EOF
				'	
				If rRs("DESCCD") ="A1" Then
						A1Desc=rRs("Description")
						A1MC1=CDbl(rRs("ACT_MTD_BAL"))
						A1MC2=CDbl(rRs("BGT_MTD_BAL"))
						A1MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						A1YC1=CDbl(rRs("ACT_YTD_BAL"))
						A1YC2=CDbl(rRs("BGT_YTD_BAL"))
						A1YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 

				If rRs("DESCCD") ="A5" Then
						A5Desc=rRs("Description")
						A5MC1=CDbl(rRs("ACT_MTD_BAL"))
						A5MC2=CDbl(rRs("BGT_MTD_BAL"))
						A5MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						A5YC1=CDbl(rRs("ACT_YTD_BAL"))
						A5YC2=CDbl(rRs("BGT_YTD_BAL"))
						A5YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 

				If rRs("DESCCD") ="B1" Then
						B1Desc=rRs("Description")
						B1MC1=CDbl(rRs("ACT_MTD_BAL"))
						B1MC2=CDbl(rRs("BGT_MTD_BAL"))
						B1MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						B1YC1=CDbl(rRs("ACT_YTD_BAL"))
						B1YC2=CDbl(rRs("BGT_YTD_BAL"))
						B1YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 

				If rRs("DESCCD") ="B5" Then
						B5Desc=rRs("Description")
						B5MC1=CDbl(rRs("ACT_MTD_BAL"))
						B5MC2=CDbl(rRs("BGT_MTD_BAL"))
						B5MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						B5YC1=CDbl(rRs("ACT_YTD_BAL"))
						B5YC2=CDbl(rRs("BGT_YTD_BAL"))
						B5YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 

				If rRs("DESCCD") ="C1" Then
						C1Desc=rRs("Description")
						C1MC1=CDbl(rRs("ACT_MTD_BAL"))
						C1MC2=CDbl(rRs("BGT_MTD_BAL"))
						C1MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						C1YC1=CDbl(rRs("ACT_YTD_BAL"))
						C1YC2=CDbl(rRs("BGT_YTD_BAL"))
						C1YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 

				If rRs("DESCCD") ="C5" Then
						C5Desc=rRs("Description")
						C5MC1=CDbl(rRs("ACT_MTD_BAL"))
						C5MC2=CDbl(rRs("BGT_MTD_BAL"))
						C5MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						C5YC1=CDbl(rRs("ACT_YTD_BAL"))
						C5YC2=CDbl(rRs("BGT_YTD_BAL"))
						C5YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 

				If rRs("DESCCD") ="E5" Then
						E5Desc="Investment Income"
						E5MC1=CDbl(rRs("ACT_MTD_BAL"))
						E5MC2=CDbl(rRs("BGT_MTD_BAL"))
						E5MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						E5YC1=CDbl(rRs("ACT_YTD_BAL"))
						E5YC2=CDbl(rRs("BGT_YTD_BAL"))
						E5YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 

				If rRs("DESCCD") ="K5" Then
						K5Desc="PBT - Company"
						K5MC1=CDbl(rRs("ACT_MTD_BAL"))
						K5MC2=CDbl(rRs("BGT_MTD_BAL"))
						K5MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						K5YC1=CDbl(rRs("ACT_YTD_BAL"))
						K5YC2=CDbl(rRs("BGT_YTD_BAL"))
						K5YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 


				If rRs("DESCCD") ="R5" Then
						R5Desc="PBT - Operations"
						R5MC1=CDbl(rRs("ACT_MTD_BAL"))
						R5MC2=CDbl(rRs("BGT_MTD_BAL"))
						R5MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))
						R5YC1=CDbl(rRs("ACT_YTD_BAL"))
						R5YC2=CDbl(rRs("BGT_YTD_BAL"))
						R5YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))
				End If 


				rRs.MoveNext
				
			Loop
			
			'
			Tbl1 = ""

			Tbl1 = Tbl1 & "<Table border=0  width='100%' cellPadding=2 cellspacing=0>"
			
			Tbl1 = Tbl1 & "<TR Class=AllReportHeader><TD class='border-ALL' Align='center' colspan=10>Month " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & " </td></TR>"	


			Tbl1 = Tbl1 & "<TR Class=RepfontStylePage14><TD class='no-top-right-border' Align='center' valign='center' >Description</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' colspan=2>vs Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Variance<br>% vs<br>Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Last year<br>(LY)</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-border' Align='center' >Variance<br>vs LY</td>"
			Tbl1 = Tbl1 & "</tr>"
			
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & A1Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A1MC1<>0 then
				Tbl1 = Tbl1 & formatnumber(A1MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"

			If A1MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(A1MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"

			If (A1MC1 - A1MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (A1MC1 - A1MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (A1MC1 - A1MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"

			If A1MC1 - A1MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(A1MC1 - A1MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			
			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A1MC2 <> 0 then
				If ((A1MC1*100)/A1MC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((A1MC1*100)/A1MC2)-100,1,,(-1)) & "% "	
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"

			If A1MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(A1MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			 
			Tbl1 = Tbl1 & "</TD>"

			If (A1MC1 - A1MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=right border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (A1MC1 - A1MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (A1MC1 - A1MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If A1MC1 - A1MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(A1MC1 - A1MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If A1MC3 <> 0 then
				If cdbl(((A1MC1*100)/A1MC3)-100) <> 0 Then
					If cdbl(((A1MC1*100)/A1MC3)-100) >= 1000 Then
						Tbl1 = Tbl1 & ">1000%"
					ElseIf cdbl(((A1MC1*100)/A1MC3)-100) <= -1000 Then
						Tbl1 = Tbl1 & ">(1000%)" 
					Else
						Tbl1 = Tbl1 & formatnumber(((A1MC1*100)/A1MC3)-100,1,,(-1)) & "%"
					End if
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"


'''''A5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & A5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"

			If A5MC1<>0 then
				Tbl1 = Tbl1 & formatnumber(A5MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			
			
			
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A5MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(A5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			If (A5MC1 - A5MC2)= 0 then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (A5MC1 - A5MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (A5MC1 - A5MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 

			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If A5MC1 - A5MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(A5MC1 - A5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			
			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A5MC2 <> 0 then
				If ((A5MC1*100)/A5MC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((A5MC1*100)/A5MC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"

			If A5MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(A5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			
			Tbl1 = Tbl1 & "</TD>"


			If (A5MC1 - A5MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (A5MC1 - A5MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (A5MC1 - A5MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If A5MC1 - A5MC3 <> 0 then
							Tbl1 = Tbl1 & formatnumber(A5MC1 - A5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If A5MC3 <> 0 then
				If ((A5MC1*100)/A5MC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((A5MC1*100)/A5MC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''A5 End
'''''B1 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & B1Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B1MC1<>0 then
				Tbl1 = Tbl1 & formatnumber(B1MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B1MC2 <>0 then
				Tbl1 = Tbl1 & formatnumber(B1MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"

		
			If (B1MC1 - B1MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (B1MC1 - B1MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (B1MC1 - B1MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If B1MC1 - B1MC2 <>0 then
				Tbl1 = Tbl1 & formatnumber(B1MC1 - B1MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B1MC2 <> 0 then
				If ((B1MC1*100)/B1MC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((B1MC1*100)/B1MC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B1MC3 <>0 then
				Tbl1 = Tbl1 & formatnumber(B1MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"

			If (B1MC1 - B1MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (B1MC1 - B1MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (B1MC1 - B1MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"

			If B1MC1 - B1MC3 <>0 then
				Tbl1 = Tbl1 & formatnumber(B1MC1 - B1MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If B1MC3 <> 0 then
				If ((B1MC1*100)/B1MC3)-100 <> 0 Then
				
					If ((B1MC1*100)/B1MC3)-100 >= 1000 Then
						Tbl1 = Tbl1 & ">1000%"
					ElseIf ((B1MC1*100)/B1MC3)-100 <= -1000 Then
						Tbl1 = Tbl1 & ">(1000%)" 
					Else
						Tbl1 = Tbl1 & formatnumber(((B1MC1*100)/B1MC3)-100,1,,(-1)) & "%"
					End If
					
					
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''B1 End

		

'''''B5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & B5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B5MC1 <>0 then
				Tbl1 = Tbl1 & formatnumber(B5MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B5MC2 <>0 then
				Tbl1 = Tbl1 & formatnumber(B5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"

			If (B5MC1 - B5MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (B5MC1 - B5MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (B5MC1 - B5MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If B5MC1 - B5MC2 <>0 then
				Tbl1 = Tbl1 & formatnumber(B5MC1 - B5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B5MC2 <> 0 then
				If ((B5MC1*100)/B5MC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((B5MC1*100)/B5MC2)-100,1,,(-1)) & "% "
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B5MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(B5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (B5MC1 - B5MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (B5MC1 - B5MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (B5MC1 - B5MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 

			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			Tbl1 = Tbl1 & formatnumber(B5MC1 - B5MC3,0,,(-1)) 
			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If B5MC3 <> 0 then
				If ((B5MC1*100)/B5MC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((B5MC1*100)/B5MC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
				
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''B5 End

'''''C1 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & C1Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If C1MC1<>0 then
			Tbl1 = Tbl1 & formatnumber(C1MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If C1MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(C1MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"


			If (C1MC1 - C1MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (C1MC1 - C1MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (C1MC1 - C1MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If C1MC1 - C1MC2 <>0 then
				Tbl1 = Tbl1 & formatnumber(C1MC1 - C1MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If C1MC2 <> 0 then
				If ((C1MC1*100)/C1MC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((C1MC1*100)/C1MC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If C1MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(C1MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (C1MC1 - C1MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (C1MC1 - C1MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (C1MC1 - C1MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 

			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"

			If C1MC1 - C1MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(C1MC1 - C1MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			
			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If C1MC3 <> 0 then
				If ((C1MC1*100)/C1MC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((C1MC1*100)/C1MC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''C1 End

'''''C5 Start
'			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
'			Tbl1 = Tbl1 & C5Desc
'			Tbl1 = Tbl1 & "</TD>"
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
'			If C5MC1<>0 then
'				Tbl1 = Tbl1 & formatnumber(C5MC1,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'	
'			Tbl1 = Tbl1 & "</TD>"
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
'			If C5MC2<>0 then
'				Tbl1 = Tbl1 & formatnumber(C5MC2,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'			Tbl1 = Tbl1 & "</TD>"
'
'			If (C5MC1 - C5MC2)= 0 then
'				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
'			ElseIf (C5MC1 - C5MC2) < 0 Then
'				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
'			ElseIf (C5MC1 - C5MC2) > 0 Then
'				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
'			End If 
'
'
'			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
'			If C5MC1 - C5MC2 <>0 then
'				Tbl1 = Tbl1 & formatnumber(C5MC1 - C5MC2,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'
'
'			Tbl1 = Tbl1 & "</TD>"
'
'
'
'
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
'			If C5MC2 <> 0 then
'				If ((C5MC1*100)/C5MC2)-100 <> 0 then
'					Tbl1 = Tbl1 & formatnumber(((C5MC1*100)/C5MC2)-100,1,,(-1)) & "%"
'				Else
'					Tbl1 = Tbl1 & "&nbsp;"
'				End If 
'
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'			Tbl1 = Tbl1 & "</TD>"
'
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
'			If C5MC3<>0 then
'				Tbl1 = Tbl1 & formatnumber(C5MC3,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'
'
'			Tbl1 = Tbl1 & "</TD>"
'
'			If (C5MC1 - C5MC3)= 0 then
'				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
'			ElseIf (C5MC1 - C5MC3) < 0 Then
'				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
'			ElseIf (C5MC1 - C5MC3) > 0 Then
'				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
'			End If 
'
'			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
'			If C5MC1 - C5MC3<>0 then
'				Tbl1 = Tbl1 & formatnumber(C5MC1 - C5MC3,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'
'
'			Tbl1 = Tbl1 & "</TD>"
'
'
'			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
'			If C5MC3 <> 0 then
'				If ((C5MC1*100)/C5MC3)-100 <> 0 then
'					Tbl1 = Tbl1 & formatnumber(((C5MC1*100)/C5MC3)-100,1,,(-1)) & "%"
'				Else
'					Tbl1 = Tbl1 & "&nbsp;"
'				End If 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'			Tbl1 = Tbl1 & "</TD>"
'			Tbl1 = Tbl1 & "</Tr>"
'
'''''C5 End

'''''R5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & R5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If R5MC1<>0 then
				Tbl1 = Tbl1 & formatnumber(R5MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If R5MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(R5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"

			If (R5MC1 - R5MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (R5MC1 - R5MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (R5MC1 - R5MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If R5MC1 - R5MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(R5MC1 - R5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If R5MC2 <> 0 then

									
				
				If ((R5MC1*100)/R5MC2)-100 <> 0 Then
						' following check added on 28-Oct-2020
						If ((R5MC1*100)/R5MC2)-100 >= 1000 Then
								Tbl1 = Tbl1 & ">1000%"
						ElseIf ((R5MC1*100)/R5MC2)-100 <= -1000 Then
							Tbl1 = Tbl1 & ">(1000%)" 
						else
							Tbl1 = Tbl1 & formatnumber(((R5MC1*100)/R5MC2)-100,1,,(-1)) & "%"
						End If 
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If R5MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(R5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"


			If (R5MC1 - R5MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (R5MC1 - R5MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (R5MC1 - R5MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If R5MC1 - R5MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(R5MC1 - R5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
			If R5MC3 <> 0 then
				If ((R5MC1*100)/R5MC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((R5MC1*100)/R5MC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''R5 End			


'''''E5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & E5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"

			If E5MC1<>0 then
				Tbl1 = Tbl1 & formatnumber(E5MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If E5MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(E5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			If (E5MC1 - E5MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (E5MC1 - E5MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (E5MC1 - E5MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If E5MC1 - E5MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(E5MC1 - E5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If E5MC2 <> 0 Then
			
				If ((E5MC1*100)/E5MC2)-100 <> 0 then
					'Tbl1 = Tbl1 & formatnumber(((E5MC1*100)/E5MC2)-100,1,,(-1)) & "%"

					If cdbl(((E5MC1*100)/E5MC2)-100) >= 1000 Then
						Tbl1 = Tbl1 & ">1000%"
					ElseIf CDbl(((E5MC1*100)/E5MC2)-100) <= -1000 Then
						Tbl1 = Tbl1 & ">(1000%)" 
					Else
						Tbl1 = Tbl1 & formatnumber((((E5MC1*100)/E5MC2)-100),1,,(-1)) & "%"
					End if


				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If E5MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(E5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			If (E5MC1 - E5MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png'></Td>"
			ElseIf (E5MC1 - E5MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (E5MC1 - E5MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If E5MC1 - E5MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(E5MC1 - E5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
			If E5MC3 <> 0 then
				If ((E5MC1*100)/E5MC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((E5MC1*100)/E5MC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''E5 End

'''''K5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' >"
			Tbl1 = Tbl1 & K5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If K5MC1<>0 then
				Tbl1 = Tbl1 & formatnumber(K5MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If K5MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(K5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (K5MC1 - K5MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (K5MC1 - K5MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (K5MC1 - K5MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 

			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If K5MC1 - K5MC2<>0 then
				Tbl1 = Tbl1 & formatnumber(K5MC1 - K5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If K5MC2 <> 0 then
				If ((K5MC1*100)/K5MC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((K5MC1*100)/K5MC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If K5MC3<>0 then
				Tbl1 = Tbl1 & formatnumber(K5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (K5MC1 - K5MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (K5MC1 - K5MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (K5MC1 - K5MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If K5MC1 - K5MC3 <>0 then
				Tbl1 = Tbl1 & formatnumber(K5MC1 - K5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
			If K5MC3 <> 0 then
				If ((K5MC1*100)/K5MC3)-100 <> 0 Then
						' This > 1000 logic was done on 31-Jan-2021				
						If ((K5MC1*100)/K5MC3)-100 >= 1000 Then
								Tbl1 = Tbl1 & ">1000%"
						ElseIf ((K5MC1*100)/K5MC3)-100 <= -1000 Then
							Tbl1 = Tbl1 & ">(1000%)" 
						else
							Tbl1 = Tbl1 & formatnumber(((K5MC1*100)/K5MC3)-100,1,,(-1)) & "%"
						End If 
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''K5 End
			Tbl1 = Tbl1 & "</Table>"

			'		
			Response.write Tbl1
			Response.write "<BR>"

'''ZTC Year

			'
			Tbl1 = ""

			Tbl1 = Tbl1 & "<Table border=0  width='100%' cellPadding=2 cellspacing=0>"
			
			Tbl1 = Tbl1 & "<TR Class=AllReportHeader><TD class='border-ALL' Align='center' colspan=10>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & " </td></TR>"	


			Tbl1 = Tbl1 & "<TR Class=RepfontStylePage14><TD class='no-top-right-border' Align='center' valign='center' >Description</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' colspan=2>vs Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Variance<br>% vs<br>Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Last year<br>(LY)</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-border' Align='center' >Variance<br>vs LY</td>"
			Tbl1 = Tbl1 & "</tr>"
			
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & A1Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A1YC1<>0 then
				Tbl1 = Tbl1 & formatnumber(A1YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A1YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(A1YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (A1YC1 - A1YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (A1YC1 - A1YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If A1YC1 - A1YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(A1YC1 - A1YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"



			' following changed on 30-Sept-2024
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right nowrap>"
			If A1YC2 <> 0 then
				If ((A1YC1*100)/A1YC2)-100 <> 0 then
				'	Tbl1 = Tbl1 & formatnumber(((A1YC1*100)/A1YC2)-100,1,,(-1)) & "%"
					If cdbl(((A1YC1*100)/A1YC2)-100) >= 1000 Then
						Tbl1 = Tbl1 & ">1000%"
					ElseIf CDbl(((A1YC1*100)/A1YC2)-100) <= -1000 Then
						Tbl1 = Tbl1 & ">(1000%)" 
					Else
						Tbl1 = Tbl1 & formatnumber(((A1YC1*100)/A1YC2)-100,1,,(-1)) & "%"
					End if
				'

				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A1YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(A1YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (A1YC1 - A1YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (A1YC1 - A1YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If A1YC1 - A1YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(A1YC1 - A1YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 


			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If A1YC3 <> 0 then
				If ((A1YC1*100)/A1YC3)-100 <> 0 then

					If ((A1YC1*100)/A1YC3)-100 >= 1000 Then
						Tbl1 = Tbl1 & ">1000%"
					ElseIf ((A1YC1*100)/A1YC3)-100 <= -1000 Then
						Tbl1 = Tbl1 & ">(1000%)" 
					Else
						Tbl1 = Tbl1 & formatnumber(((A1YC1*100)/A1YC3)-100,1,,(-1)) & "%"
					End if
					
					
					
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"


'''''A5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & A5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A5YC1 <>0 then
				Tbl1 = Tbl1 & formatnumber(A5YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A5YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(A5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (A5YC1 - A5YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (A5YC1 - A5YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If A5YC1 - A5YC2 <>0 then
				Tbl1 = Tbl1 & formatnumber(A5YC1 - A5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A5YC2 <> 0 then
				If ((A5YC1*100)/A5YC2)-100 <> 0 then
				
					If ((A5YC1*100)/A5YC2)-100 >= 1000 Then
						Tbl1 = Tbl1 & ">1000%"
					ElseIf ((A5YC1*100)/A5YC2)-100 <= -1000 Then
						Tbl1 = Tbl1 & ">(1000%)" 
					Else
						Tbl1 = Tbl1 & formatnumber(((A5YC1*100)/A5YC2)-100,1,,(-1)) & "%"
					End if
					
					
					
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If A5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(A5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (A5YC1 - A5YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (A5YC1 - A5YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If A1MCA5YC1 - A5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(A5YC1 - A5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If A5YC3 <> 0 then
				If ((A5YC1*100)/A5YC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((A5YC1*100)/A5YC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''A5 End
'''''B1 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & B1Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B1YC1<>0 then
				Tbl1 = Tbl1 & formatnumber(B1YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B1YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(B1YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (B1YC1 - B1YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (B1YC1 - B1YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If B1YC1 - B1YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(B1YC1 - B1YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B1YC2 <> 0 then
				If ((B1YC1*100)/B1YC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((B1YC1*100)/B1YC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B1YC3 <>0 then
				Tbl1 = Tbl1 & formatnumber(B1YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (B1YC1 - B1YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (B1YC1 - B1YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If B1YC1 - B1YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(B1YC1 - B1YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If B1YC3 <> 0 then
				If ((B1YC1*100)/B1YC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((B1YC1*100)/B1YC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''B1 End

		

'''''B5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' nowrap>"
			Tbl1 = Tbl1 & B5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B5YC1<>0 then
			Tbl1 = Tbl1 & formatnumber(B5YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B5YC2<>0 then
			Tbl1 = Tbl1 & formatnumber(B5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (B5YC1 - B5YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (B5YC1 - B5YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If B5YC1 - B5YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(B5YC1 - B5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B5YC2 <> 0 then
				If ((B5YC1*100)/B5YC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((B5YC1*100)/B5YC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If B5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(B5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (B5YC1 - B5YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (B5YC1 - B5YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If B5YC1 - B5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(B5YC1 - B5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If B5YC3 <> 0 then
				If ((B5YC1*100)/B5YC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((B5YC1*100)/B5YC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''B5 End

'''''C1 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & C1Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If C1YC1<>0 then
				Tbl1 = Tbl1 & formatnumber(C1YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If C1YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(C1YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (C1YC1 - C1YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (C1YC1 - C1YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If C1YC1 - C1YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(C1YC1 - C1YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"

			If C1YC2 <> 0 then
				If cdbl(((C1YC1*100)/C1YC2)-100) >= 1000 Then
					Tbl1 = Tbl1 & ">1000%"
				ElseIf CDbl(((C1YC1*100)/C1YC2)-100) <= -1000 Then
					Tbl1 = Tbl1 & ">(1000%)" 
				Else
'					Tbl1 = Tbl1 & cdbl(((C1YC1*100)/C1YC2)-100) ' formatnumber(((C1YC1*100)/C1YC2)-100,1,,(-1)) & "%"
					Tbl1 = Tbl1 & formatnumber(((C1YC1*100)/C1YC2)-100,1,,(-1)) & "%"    ' This line was modified on 29-Sept-2020
				End if

'				If ((C1YC1*100)/C1YC2)-100 <> 0 then
'					Tbl1 = Tbl1 & formatnumber(((C1YC1*100)/C1YC2)-100,1,,(-1)) & "%"
'				Else
'					Tbl1 = Tbl1 & "&nbsp;"
'				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If C1YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(C1YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (C1YC1 - C1YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (C1YC1 - C1YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If C1YC1 - C1YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(C1YC1 - C1YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If C1YC3 <> 0 then
				If ((C1YC1*100)/C1YC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((C1YC1*100)/C1YC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''C1 End

'''''C5 Start
'			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
'			Tbl1 = Tbl1 & C5Desc
'			Tbl1 = Tbl1 & "</TD>"
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
'			If C5YC1<>0 then
'				Tbl1 = Tbl1 & formatnumber(C5YC1,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'
'			Tbl1 = Tbl1 & "</TD>"
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
'			If C5YC2<>0 then
'			Tbl1 = Tbl1 & formatnumber(C5YC2,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'
'			Tbl1 = Tbl1 & "</TD>"
'
'			If (C5YC1 - C5YC2) < 0 Then
'				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
'			ElseIf (C5YC1 - C5YC2) > 0 Then
'				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
'			Else
'				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
'			End If 
'
'
'			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
'			If C5YC1 - C5YC2<>0 then
'				Tbl1 = Tbl1 & formatnumber(C5YC1 - C5YC2,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'
'			Tbl1 = Tbl1 & "</TD>"
'
'
'
'
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
'			If C5YC2 <> 0 then
'				If ((C5YC1*100)/C5YC2)-100 <> 0 then
'					Tbl1 = Tbl1 & formatnumber(((C5YC1*100)/C5YC2)-100,1,,(-1)) & "%"
'				Else
'					Tbl1 = Tbl1 & "&nbsp;"
'				End If 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'			Tbl1 = Tbl1 & "</TD>"
'
'			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
'			If C5YC3<>0 then
'				Tbl1 = Tbl1 & formatnumber(C5YC3,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'
'			Tbl1 = Tbl1 & "</TD>"
'
'			If (C5YC1 - C5YC3) < 0 Then
'				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
'			ElseIf (C5YC1 - C5YC3) > 0 Then
'				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
'			Else
'				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
'			End If 
'
'
'			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
'			If C5YC1 - C5YC3<>0 then
'			Tbl1 = Tbl1 & formatnumber(C5YC1 - C5YC3,0,,(-1)) 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'
'			Tbl1 = Tbl1 & "</TD>"
'
'
'			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
'			If C5YC3 <> 0 then
'				If ((C5YC1*100)/C5YC3)-100 <> 0 then
'					Tbl1 = Tbl1 & formatnumber(((C5YC1*100)/C5YC3)-100,1,,(-1)) & "%"
'				Else
'					Tbl1 = Tbl1 & "&nbsp;"
'				End If 
'			Else
'				Tbl1 = Tbl1 & "&nbsp;"
'			End If 
'			Tbl1 = Tbl1 & "</TD>"
'			Tbl1 = Tbl1 & "</Tr>"
'
'''''C5 End

'''''R5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & R5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If R5YC1 <>0 then
				Tbl1 = Tbl1 & formatnumber(R5YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If R5YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(R5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (R5YC1 - R5YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (R5YC1 - R5YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"

			If R5YC1 - R5YC2 <>0 then
				Tbl1 = Tbl1 & formatnumber(R5YC1 - R5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If R5YC2 <> 0 then
				If ((R5YC1*100)/R5YC2)-100-100 <> 0 Then
				
					If ((R5YC1*100)/R5YC2)-100 >= 1000 Then
						Tbl1 = Tbl1 & ">1000%"
					ElseIf ((R5YC1*100)/R5YC2)-100 <= -1000 Then
						Tbl1 = Tbl1 & ">(1000%)" 
					Else
						Tbl1 = Tbl1 & formatnumber(((R5YC1*100)/R5YC2)-100,1,,(-1)) & "%"
					End if

					
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"

			If R5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(R5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (R5YC1 - R5YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (R5YC1 - R5YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If R5YC1 - R5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(R5YC1 - R5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
			If R5YC3 <> 0 then
				If ((R5YC1*100)/R5YC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((R5YC1*100)/R5YC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''R5 End			


'''''E5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			Tbl1 = Tbl1 & E5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If E5YC1<>0 then
				Tbl1 = Tbl1 & formatnumber(E5YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If E5YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(E5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (E5YC1 - E5YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (E5YC1 - E5YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If E5YC1 - E5YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(E5YC1 - E5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If E5YC2 <> 0 then
				If ((E5YC1*100)/E5YC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((E5YC1*100)/E5YC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If E5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(E5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (E5YC1 - E5YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (E5YC1 - E5YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If E5YC1 - E5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(E5YC1 - E5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
			If E5YC3 <> 0 then
				If ((E5YC1*100)/E5YC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((E5YC1*100)/E5YC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''E5 End

'''''K5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' >"
			Tbl1 = Tbl1 & K5Desc
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If K5YC1 <>0 then
				Tbl1 = Tbl1 & formatnumber(K5YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If K5YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(K5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			If (K5YC1 - K5YC2) = 0 Then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (K5YC1 - K5YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (K5YC1 - K5YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 



			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If K5YC1 - K5YC2<>0 then
				Tbl1 = Tbl1 & formatnumber(K5YC1 - K5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If K5YC2 <> 0 then
				If ((K5YC1*100)/K5YC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((K5YC1*100)/K5YC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If K5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(K5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"

			If (K5YC1 - K5YC3) = 0 Then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (K5YC1 - K5YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (K5YC1 - K5YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If K5YC1 - K5YC3<>0 then
				Tbl1 = Tbl1 & formatnumber(K5YC1 - K5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
			If K5YC3 <> 0 Then
				If ((K5YC1*100)/K5YC3)-100 <> 0 then

					If ((K5YC1*100)/K5YC3)-100 >= 1000 Then
						Tbl1 = Tbl1 & ">1000%"
					ElseIf ((K5YC1*100)/K5YC3)-100 <= -1000 Then
						Tbl1 = Tbl1 & ">(1000%)" 
					Else
						Tbl1 = Tbl1 & formatnumber(((K5YC1*100)/K5YC3)-100,1,,(-1)) & "%"
					End if

					
					
				Else
					Tbl1 = Tbl1 & "&nbsp;"
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;"
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''K5 End
			Tbl1 = Tbl1 & "</Table>"

			'		
			Response.write Tbl1


		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched For ZTC From Deptwise PBT **********</Td>"
		End if		
		rRs.close

	


''''''''''''''''For OFCC DEPTWISE PBT Start
		sSql="" 

		sSql= "SELECT  DESCCD, DESCRIPTION, ACT_MTD_BAL , BGT_MTD_BAL,ACT_YTD_BAL , BGT_YTD_BAL,ACT_MTD_BAL_PY1,ACT_YTD_BAL_PY1" 
		sSql= sSql & " FROM MISTXN   " 
		sSql= sSql & " where Trim(cMonth)='" & TheReportMonth & "'  " 
		sSql= sSql & " AND Trim(cYear)='" & TheReportYear & "'  " 
		sSql= sSql & " AND ccompcode='OFCC' and desccd in ('M5','R5') " 
		sSql= sSql & " order by 1 "

'response.write sSql
'response.end
		rRs.Open Trim(sSql),cCn


		If Not rRs.Eof Then
			'
			M5Desc=""
			M5MC1=0
			M5MC2=0
			M5MC3=0
			M5YC1=0
			M5YC2=0
			M5YC3=0

			R5Desc=""
			R5MC1=0
			R5MC2=0
			R5MC3=0
			R5YC1=0
			R5YC2=0
			R5YC3=0
			Do while not rRs.EOF
				'	
				If rRs("DESCCD") ="M5" Then
						M5Desc=rRs("Description")
						M5MC1=CDbl(rRs("ACT_MTD_BAL"))/2
						M5MC2=CDbl(rRs("BGT_MTD_BAL"))/2
						M5MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))/2
						M5YC1=CDbl(rRs("ACT_YTD_BAL"))/2
						M5YC2=CDbl(rRs("BGT_YTD_BAL"))/2
						M5YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))/2
				End If 


				If rRs("DESCCD") ="R5" Then
						R5Desc=rRs("Description")
						R5MC1=CDbl(rRs("ACT_MTD_BAL"))/2
						R5MC2=CDbl(rRs("BGT_MTD_BAL"))/2
						R5MC3=CDbl(rRs("ACT_MTD_BAL_PY1"))/2
						R5YC1=CDbl(rRs("ACT_YTD_BAL"))/2
						R5YC2=CDbl(rRs("BGT_YTD_BAL"))/2
						R5YC3=CDbl(rRs("ACT_YTD_BAL_PY1"))/2
				End If 

				rRs.MoveNext
				
			Loop
			
			'
		
			sSql =  "<br><Table border=0  width='95%' cellPadding=2 cellspacing=0	>"
			sSql = sSql & "<TR Class=GreenTransparentHeading><TD >" 
			sSql = sSql & "Oman Formaldehyde Chemical Company (OFCC)" 
			sSql = sSql & "</Td >" 
			sSql = sSql & "</TR>" 
			sSql = sSql & "</Table>"
			response.write sSql	

			Tbl1 = ""

			Tbl1 = Tbl1 & "<Table border=0  width='100%' cellPadding=1 cellspacing=0>"
			
			Tbl1 = Tbl1 & "<TR Class=AllReportHeader><TD class='border-ALL' Align='center' colspan=10>Month " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & " </td></TR>"	


			'Tbl1 = Tbl1 & "<TR Class=RepfontStyle><TD class='no-top-right-border' Align='center' valign='center' >Description</td>"	
			'Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Flash</td>"	
			'Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Budget</td>"	
			'Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' colspan=2>vs Budget</td>"	
			'Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' >Variance<br>% vs<br>Budget</td>"	
			'Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='59'>Last year<br>(LY)</td>"	
			'Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' colspan=2>vs LY</td>"	
			'Tbl1 = Tbl1 & "<TD class='no-top-border' Align='center' >Variance<br>vs LY</td>"
			'Tbl1 = Tbl1 & "</tr>"
			
			Tbl1 = Tbl1 & "<TR Class=RepfontStylePage14><TD class='no-top-right-border' Align='center' valign='center' width='231' >Description</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='49' >Flash</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='54' >Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' colspan=2 width='69'>vs Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='71' >Variance<br>% vs<br>Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='59'>Last year<br>(LY)</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='59' colspan=2>vs LY</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-border' Align='center' width='59'>Variance<br>vs LY</td>"
			Tbl1 = Tbl1 & "</tr>"



'''''R5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			'Tbl1 = Tbl1 & R5Desc & " (50%)"
			Tbl1 = Tbl1 & "PBT - Operations (50%)"
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If R5MC1 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If R5MC2 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"

			If (R5MC1 - R5MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (R5MC1 - R5MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (R5MC1 - R5MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If R5MC1 - R5MC2 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5MC1 - R5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If R5MC2 <> 0 Then
				If ((R5MC1*100)/R5MC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((R5MC1*100)/R5MC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;" 
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If R5MC3 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"


			If (R5MC1 - R5MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			ElseIf (R5MC1 - R5MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (R5MC1 - R5MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If R5MC1 - R5MC3 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5MC1 - R5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If R5MC3 <> 0 then
				If ((R5MC1*100)/R5MC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((R5MC1*100)/R5MC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;" 
				End If 
				
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''R5 End			

'''''M5 Start

			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' >"
			'Tbl1 = Tbl1 & M5Desc & " (50%)"
			Tbl1 = Tbl1 & "Cash PBT - Operations (50%)"
			
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If M5MC1 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5MC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If M5MC2 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"

			If (M5MC1 - M5MC2)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (M5MC1 - M5MC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (M5MC1 - M5MC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If M5MC1 - M5MC2 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5MC1 - M5MC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If M5MC2 <> 0 then
				If ((M5MC1*100)/M5MC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((M5MC1*100)/M5MC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;" 
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If M5MC3 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"

			If (M5MC1 - M5MC3)= 0 then
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			ElseIf (M5MC1 - M5MC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (M5MC1 - M5MC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If M5MC1 - M5MC3 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5MC1 - M5MC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
			If M5MC3 <> 0 then
				If ((M5MC1*100)/M5MC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((M5MC1*100)/M5MC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;" 
				End If 
			Else
				Tbl1 = Tbl1 & "0%" 
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"
'''''M5 End
			Tbl1 = Tbl1 & "</Table>"

			'		
			Response.write Tbl1
			Response.write "<BR>"

'''OFCC Year

			'
			Tbl1 = ""

			Tbl1 = Tbl1 & "<Table border=0  width='100%' cellPadding=2 cellspacing=0>"
			
			Tbl1 = Tbl1 & "<TR Class=AllReportHeader><TD class='border-ALL' Align='center' colspan=10>YTD " & FullMonthsArr(cint(TheReportMonth)-1) & "&nbsp;" & TheReportYear & " </td></TR>"	


			Tbl1 = Tbl1 & "<TR Class=RepfontStylePage14><TD class='no-top-right-border' Align='center' valign='center' width='231' >Description</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='49' >Flash</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='54' >Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' colspan=2 width='69'>vs Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='71' >Variance<br>% vs<br>Budget</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='59'>Last year<br>(LY)</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' Align='center' width='59' colspan=2>vs LY</td>"	
			Tbl1 = Tbl1 & "<TD class='no-top-border' Align='center' width='59'>Variance<br>vs LY</td>"
			Tbl1 = Tbl1 & "</tr>"
			


'''''R5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' >"
			'Tbl1 = Tbl1 & R5Desc & " (50%)"
			Tbl1 = Tbl1 & "PBT - Operations (50%)"
			
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If R5YC1 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If R5YC2 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"

			If (R5YC1 - R5YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (R5YC1 - R5YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If R5YC1 - R5YC2 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5YC1 - R5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If R5YC2 <> 0 then
				If ((R5YC1*100)/R5YC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((R5YC1*100)/R5YC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;" 
				End If 

			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border_Grey' align=right>"
			If R5YC3 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"

			If (R5YC1 - R5YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/down.png' ></Td>"
			ElseIf (R5YC1 - R5YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border_Grey'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border_Grey' align=right>"
			If R5YC1 - R5YC3 <> 0 then
				Tbl1 = Tbl1 & formatnumber(R5YC1 - R5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border_Grey' align=right>"
			If R5YC3 <> 0 then
				If ((R5YC1*100)/R5YC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((R5YC1*100)/R5YC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;" 
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''R5 End			

'''''''''M5 Start
			Tbl1 = Tbl1 & "<Tr Class=RepContentFontPage14>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' >"
			'Tbl1 = Tbl1 & M5Desc & "**** (50%)"
			Tbl1 = Tbl1 & "Cash PBT - Operations (50%)"
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If M5YC1 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5YC1,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If M5YC2 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"

			If (M5YC1 - M5YC2) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (M5YC1 - M5YC2) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If M5YC1 - M5YC2 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5YC1 - M5YC2,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"




			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			If M5YC2 <> 0 then
				If ((M5YC1*100)/M5YC2)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((M5YC1*100)/M5YC2)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;" 
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End If 
			Tbl1 = Tbl1 & "</TD>"

			Tbl1 = Tbl1 & "<TD class='no-top-right-border' align=right>"
			Tbl1 = Tbl1 & formatnumber(M5YC3,0,,(-1)) 
			Tbl1 = Tbl1 & "</TD>"

			If (M5YC1 - M5YC3) < 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/down.png' ></Td>"
			ElseIf (M5YC1 - M5YC3) > 0 Then
				Tbl1 = Tbl1 & "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/up.png' ></Td>"
			Else
				Tbl1 = Tbl1 &  "<TD Align=center border=0 width='30' class='no-top-right-border'><img src='images/zero.png' ></Td>"
			End If 


			Tbl1 = Tbl1 & "<TD class='no-right-left-top-border' align=right>"
			If M5YC1 - M5YC3 <> 0 then
				Tbl1 = Tbl1 & formatnumber(M5YC1 - M5YC3,0,,(-1)) 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End if

			Tbl1 = Tbl1 & "</TD>"


			Tbl1 = Tbl1 & "<TD class='no-top-border' align=right>"
			If M5YC3 <> 0 then
				If ((M5YC1*100)/M5YC3)-100 <> 0 then
					Tbl1 = Tbl1 & formatnumber(((M5YC1*100)/M5YC3)-100,1,,(-1)) & "%"
				Else
					Tbl1 = Tbl1 & "&nbsp;" 
				End If 
			Else
				Tbl1 = Tbl1 & "&nbsp;" 
			End If 
			Tbl1 = Tbl1 & "</TD>"
			Tbl1 = Tbl1 & "</Tr>"

'''''''''M5 End

			Tbl1 = Tbl1 & "</Table>"

			'		
			Response.write Tbl1


		Else
			Response.write "<Tr Class=cellStyle1><TD class='no-top-right-border' Align=Center Colspan=27><b>**********No Record Fetched For OFCC Flash**********</Td>"
		End if		
		rRs.close
''''''''''''''''For OFCC DEPTWISE PBT End

	End function


	Private Function IndexPrint()


			sSql = ""

			sSql = sSql & "<Table border=0  width='100%' height='75%' cellPadding=2 cellspacing=0 >"
			
			sSql = sSql & "<TR Class=IndexHead1Font><TD class='border-ALL' Align='Center' colspan=2>INDEX</td></TR>"	

			sSql = sSql & "<TR Class=IndexHead2Font>"
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Contents</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >Page No.</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Group Results - Group Level</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >1</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Revenue by Company</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >2</td>"	
			sSql = sSql & "</tr>"
		
			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >PBT-Operations by Company (excl. Investment Income)</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >3</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Cash PBT-Operations by Company (excl. Investment Income)</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >4</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Manufacturing Sector - Month</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >5</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Trading Sector & Contracting Sector - Month</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >5</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Services and Investment Sector - Month</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >5</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Manufacturing Sector - YTD</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >6</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Trading Sector & Contracting Sector - YTD</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >6</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Services and Investment Sector - YTD</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >6</td>"	
			sSql = sSql & "</tr>"


			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Major contributions to Profit/(Loss) from Operations (excl. Investment Income)</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >7</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Investment Income & Inter-Company Dividends</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >7</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >MTD Variance Analysis versus Budget - Manufacturing Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >8</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >MTD Variance Analysis versus Last Year - Manufacturing Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >8</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >MTD Variance Analysis versus Budget - Trading Sector & Contracting Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >9</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >MTD Variance Analysis versus Last Year - Trading Sector & Contracting Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >9</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >MTD Variance Analysis versus Budget - Services & Investment Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >10</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >MTD Variance Analysis versus Last Year - Services & Investment Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >10</td>"	
			sSql = sSql & "</tr>"



			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >YTD Variance Analysis versus Budget - Manufacturing Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >11</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >YTD Variance Analysis versus Last Year - Manufacturing Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >11</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >YTD Variance Analysis versus Budget - Trading Sector & Contracting Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >12</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >YTD Variance Analysis versus Last Year - Trading Sector & Contracting Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >12</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >YTD Variance Analysis versus Budget - Services & Investment Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >13</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >YTD Variance Analysis versus Last Year - Services & Investment Sector</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >13</td>"	
			sSql = sSql & "</tr>"





			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >PBT for Zawawi Trading Company & Oman Formaldehyde Chemical Company</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >14</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Company-wise Performance Summary of Group Results for the period - MTD</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >15</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Company-wise Performance Summary of Group Results for the period - YTD</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >16</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "<TR Class=IndexContentFont>"	
			sSql = sSql & "<TD class='no-top-right-border' Align='left' >Sectorwise Summary of Group Results</td>"	
			sSql = sSql & "<TD class='no-top-border' Align='center' >17</td>"	
			sSql = sSql & "</tr>"

			sSql = sSql & "</Table>"	
			Response.write sSql 




	End function


' style='padding: 0 2px 0 2px;' T L B R
Private Function Annexures

			sSql = "<Table border=0  width='100%' height='100%' cellPadding=0 cellspacing=0 >"
			sSql = sSql & "<TR Class=AnnexFont ><TD Align='Center' valign='Center'>ANNEXURES</td></tr>"	
			sSql = sSql & "</tr>"
			sSql = sSql & "</table>"

			Response.write sSql 
End function

Private Function CoverPage

			sSql = "<Table border=0  width='100%' height='100%' cellPadding=2 cellspacing=0 >"
			
			sSql = sSql & "<TR Class=IndexHead1Font height='3%'><TD Align='Center' valign='Top'><font color='red'>CONFIDENTIAL</td></tr>"	
			sSql = sSql & "<TR Class=IndexHead1Font>"	
			sSql = sSql & "<TD Align='Center' valign='Center' ><font size='+3'>GROUP RESULTS"	
			sSql = sSql & "<br><br>" & UCase(FullMonthsArr(cint(TheReportMonth)-1)) & "&nbsp;" & TheReportYear & "</font></td>"	
			sSql = sSql & "</tr>"
			sSql = sSql & "</table>"

			Response.write sSql 
End function


'	If ((R5YC1*100)/R5YC2)-100 >= 1000 Then
'		Tbl1 = Tbl1 & ">1000%"
'	ElseIf ((R5YC1*100)/R5YC2)-100 <= -1000 Then
'		Tbl1 = Tbl1 & ">(1000%)" 
'	Else
'		Tbl1 = Tbl1 & formatnumber(((R5YC1*100)/R5YC2)-100,1,,(-1)) & "%"
'	End if
'page 5 And page 14
%>

