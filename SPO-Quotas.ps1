##By Kristopher Roy AKA TankCR
##Last Modified 10May2023
##Blog URL = https://tankcr.blogspot.com/

#The First thing we must do is call our SPOCreds Script
#$SPOCred = "PathToScriptFile such as C:\Scripts\SPOCreds.PS1"
#invoke-expression -Command $SPOCred
#Set your Cloud Admin URL
$org = "GTIL"
$Url = "https://gtinetorg-admin.sharepoint.com"
Connect-SPOService -url $url

#Now Set your Output File Locations, use your temp folder

$File = "C:\TEMP\usage-04Oct23.xml"
$File2 = "C:\TEMP\usage-04Oct23.xlsx"

#As I want this to be scheduled and Faced on my On-Prem SharePoint I set a net location to put the final copy
#$NetLoc = "\\myonpremsharepoint.local\sites\reporting\Shared Documents\"

#This is our first Pattern, in this Pattern I am grabbing three different sections of the URL's so that I can sort my report automatically based on the core part that is similar across my sites. You will need to modify this to your site use http://regex101.com/ if you need help doing so...
#$Pat = "(.*)(CRM.*_?)(_.*)"

#Now we grab and store our sites in a sorted manner, I am first sorting by the pattern we set up top, then I am sorting by the sites using the highest percentage of their quotas.
#$Sites = (Get-SPOSite -Detailed -ErrorAction Ignore -limit all)|sort -Descending -Property URL,@{e={[INT]($_.StorageUsageCurrent)/[INT]($_.StorageQuota)}}
$Sites = (Get-SPOSite -Detailed -ErrorAction Ignore -limit all)|sort -Descending|select *,@{e={[INT]($_.StorageUsageCurrent)/[INT]($_.StorageQuota)}}

#The Excel formatting requires a proper count on our Rows or it will error out, so we are going to grab a count on how many sites we have:
$Number = $Sites.Count

#Lets create our XML File, this is the initial formatting that it will need to understand what it is, and what styles we are using.
(
 '<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40"
 xmlns:x2="http://schemas.microsoft.com/office/excel/2003/xml"
 xmlns:udc="http://schemas.microsoft.com/data/udc"
 xmlns:xsd="http://www.w3.org/2001/XMLSchema"
 xmlns:udcxf="http://schemas.microsoft.com/data/udc/xmlfile">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Kristopher Roy</Author>
  <LastAuthor>'+$env:USERNAME+'</LastAuthor>
  <Created>'+(get-date)+'</Created>
   <Version>15.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>12435</WindowHeight>
  <WindowWidth>25500</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>0</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s27" ss:Name="Bad">
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#800000"/>
   <Interior ss:Color="#FFC7CE" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s26" ss:Name="Good">
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#008000"/>
   <Interior ss:Color="#C6EFCE" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s62" ss:Name="Hyperlink">
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0066CC"
    ss:Underline="Single"/>
  </Style>
  <Style ss:ID="s28" ss:Name="Neutral">
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#993300"/>
   <Interior ss:Color="#FFEB9C" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s63">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
    ss:Bold="1"/>
   <Interior ss:Color="#D0CECE" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s64">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
  </Style>
  <Style ss:ID="s66" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
  </Style>
  <Style ss:ID="s68" ss:Parent="s26">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <NumberFormat ss:Format="0%"/>
  </Style>
  <Style ss:ID="s70" ss:Parent="s28">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <NumberFormat ss:Format="0%"/>
  </Style>
  <Style ss:ID="s72" ss:Parent="s27">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <NumberFormat ss:Format="0%"/>
  </Style>
    <Style ss:ID="s90" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0066CC"
    ss:Bold="1" ss:Underline="Single"/>
   <Interior ss:Color="#FFE699" ss:Pattern="Solid"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Site Quotas">
  <Table ss:ExpandedColumnCount="9" ss:ExpandedRowCount="'+($Number+1)+'" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="100.5"/>
   <Column ss:Width="441.75"/>
   <Column ss:Width="247.5"/>
   <Column ss:Width="33.75"/>
   <Column ss:Width="41.25"/>
   <Column ss:Width="73.5"/>
   <Column ss:Width="77.25"/>
   <Column ss:Width="117.75"/>
   <Column ss:Width="105"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s63"><Data ss:Type="String">Title</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="String">URL</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="String">Owner</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="String">Quota</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="String">Percent</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="String">QuotaWarning</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="String">ResourceQuota</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="String">ResourceQuotaWarning</Data></Cell>
    <Cell ss:StyleID="s63"><Data ss:Type="String">StorageUsageCurrent</Data></Cell>
   </Row>')> $file

#We are now going to iterate through each site collection and write them to our xml file. But of course we first need to set some parameters on how this will occur so that we can get the results we want.
FOREACH ($Site in $Sites){

#Gets the Site Owner, if it can't for some reason it sets the variable to just "Farm"
$Owner = if(($Site.Owner -eq $NULL) -or ($site.Owner -eq "")){$Owner = "Farm"}ELSE{$Site.Owner}

#Another Pattern, this one is to break down the URL's so that we can set the Title if one doesn't exist
$Pat = "(https?.\/\/)(\w*-?\w*)(.*)?"
if($Site.Title -eq $NULL -or $Site.Title -eq ""){$Title = ($Site.url -replace $pat,'$2')}ELSE{$Title = $Site.Title}

#Gets the Usage Percentage then sets the excel style so that higher percentages get color coded
$BasePercent = ($Site.StorageUsageCurrent / $Site.StorageQuota)
If ($BasePercent -ge ".75"){$Percent = '<Cell ss:StyleID="s72"><Data ss:Type="Number">'+$BasePercent+'</Data></Cell>'}
If (($BasePercent -lt ".75") -and ($BasePercent -ge ".60")){$Percent = '<Cell ss:StyleID="s70"><Data ss:Type="Number">'+$BasePercent+'</Data></Cell>'}
If ($BasePercent -lt ".60"){$Percent = '<Cell ss:StyleID="s68"><Data ss:Type="Number">'+$BasePercent+'</Data></Cell>'}

#I wanted any sites with Prod in them to be highlighted, you can change this to whatever you like
$URLCODE = IF($Site.URL -ilike "*Prod*"){"s90"}ELSE{"s66"}
add-content $File ('<Row ss:AutoFitHeight="0">')
add-content $File ('<Cell ss:StyleID="s64"><Data ss:Type="String">'+($Title)+'</Data></Cell>')
add-content $File ('<Cell ss:StyleID="'+($URLCODE)+'" ss:HRef="'+($Site.url)+'">'+'<Data ss:Type="String">'+($Site.url)+'</Data></Cell>')
add-content $File ('<Cell ss:StyleID="s64"><Data ss:Type="String">'+($Owner)+'</Data></Cell>')
add-content $File ('<Cell ss:StyleID="s64"><Data ss:Type="Number">'+($Site.StorageQuota)+'</Data></Cell>')
add-content $file $Percent
add-content $File ('<Cell ss:StyleID="s64"><Data ss:Type="Number">'+($Site.StorageQuotaWarningLevel)+'</Data></Cell>')
add-content $File ('<Cell ss:StyleID="s64"><Data ss:Type="Number">'+($Site.ResourceQuota)+'</Data></Cell>')
add-content $File ('<Cell ss:StyleID="s64"><Data ss:Type="Number">'+($Site.ResourceQuotaWarningLevel)+'</Data></Cell>')
add-content $File ('<Cell ss:StyleID="s64"><Data ss:Type="Number">'+($Site.StorageUsageCurrent)+'</Data></Cell>')
add-content $File '</Row>'
}

#The last thing that we will write to our XML file is an XML component that closes it out.
add-content $file (
'    </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Unsynced/>
   <Print>
    <ValidPrinterInfo/>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <LeftColumnVisible>1</LeftColumnVisible>
   <FreezePanes/>
   <FrozenNoSplit/>
   <SplitHorizontal>1</SplitHorizontal>
   <TopRowBottomPane>1</TopRowBottomPane>
   <ActivePane>2</ActivePane>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveCol>1</ActiveCol>
    </Pane>
    <Pane>
     <Number>2</Number>
     <ActiveRow>10</ActiveRow>
     <ActiveCol>2</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>')

#At this point we will grab our file and convert it to a true excel file so that you can use it with the excel viewer web-part, we will also copy it to your network location and then delete it off your hard drive. 
#(Important Note, to convert it to XSLX you must have excel installed on the machine where you run this) 
$objExcel = new-object -comobject excel.application
$UserWorkBook = $objExcel.Workbooks.Open($file)
$UserWorkBook.SaveAs($file2,51)
$UserWorkBook.Close()

#copy $file2 $NetLoc
#remove-item $file
#remove-item $file2
