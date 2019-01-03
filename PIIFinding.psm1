<#
.SYNOPSIS
    PII Searching Script
.DESCRIPTION
    Long description
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect() | Out-Null

$ConfigRoot = "PIIFinding\config"
$ReportRoot = "PIIFinding\reports"
$LogRoot = "PIIFinding\logs"

try
{
    Add-Type -AssemblyName System.windows.forms
}
catch
{
    $_.Exception.LoaderExceptions | ForEach-object{
        throw
    }
}
function Find-PDF {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)][String]$File
    )
    
    begin {
        Try { [io.file]::OpenWrite($File).close() }
        Catch { Write-Log -Message "Access Denied" -Severity 2 -File $File -FileType "PDF"; return }
    }
    
    process {
        try{
            $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file
            $fields = $reader.AcroFields.Fields

                if($fields.Count -gt 0){
                    foreach ($field in $fields.Keys){
                        [string]$FieldData = $reader.AcroFields.GetField($field.ToString())
                        $FieldData | Get-PIIString -FileType "PDF" | Where-Object{ $_.isMatch() } | ForEach-Object {  Format-PII -File $File -Match $_;break}
                    }
                }

                :pages for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
                    $TextData = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader,$page).Split([char]0x000A)
                        foreach($line in $TextData){
                            $line | Get-PIIString -FileType "PDF" | Where-Object {$_.isMatch() } | ForEach-Object {  Format-PII -File $File -Match $_;break pages}
                        }
                }
                
        }catch{
            if($_.Exception.GetType() -eq [System.UnauthorizedAccessException] ){
                exit
            }
            Write-Log -Message $_ -Severity 2 -File $File -FileType "PDF"
        }
        finally{
            $reader.Close()
        }
    }
}
Export-ModuleMember -Function Find-PDF
function Find-FlatFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [String]$File
    )
    
    begin {
        Try { [io.file]::OpenWrite($File).close() }
        Catch { Write-Log -Message "Access Denied" -Severity 3 -File $File -FileType "CSV"; return }
    }
    
    process {
        try{
            $reader = [System.IO.File]::OpenText($File)

            :lines while($null -ne ($line = $reader.ReadLine())) {
                if($line -eq ""){continue}

                $line | Get-PIIString -FileType "CSV" | Where-Object { $_.isMatch() } | ForEach-Object { Format-PII -File $File -Match $_;break lines }
            }
        }
        catch{
            if($_.Exception.GetType() -eq [System.UnauthorizedAccessException] ){
                exit
            }
            Write-Log -File $File -Message "$_ Failed!" -Severity 2 -FileType "CSV"
            
        }
        finally{
            if($null -ne $reader){
                $reader.Close()
            }
        }
    }
}
Export-ModuleMember  -Function Find-FlatFile
function Find-XLSX {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string]$File
    )
    
    begin {
        Try { 
            [io.file]::OpenWrite($File).close() 
        }
        Catch{ 
            Write-Log -Message "Access Denied" -Severity 2 -File $File -FileType "XLSX"; return 
        }
    }
    
    process {
        try{
            $spreadsheetDocument = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($File, $false)
            $workbookPart = $spreadsheetDocument.WorkbookPart
            $ShareStringTablePart = $workbookPart.SharedStringTablePart
            
            if($null -eq $ShareStringTablePart){return}

            $reader = [DocumentFormat.OpenXml.OpenXmlPartReader]::new([DocumentFormat.OpenXml.Packaging.OpenXmlPart]$ShareStringTablePart, $false)

            :lines while($reader.Read())
            {
                if ($reader.ElementType -eq [DocumentFormat.OpenXml.Spreadsheet.SharedStringItem])
                {
                    $ssi = $reader.LoadCurrentElement()
                    $value = $ssi.InnerText

                    if($null -eq $value -or $value -eq ""){continue}
    
                    $value | Get-PIIString -FileType "XLSX" | Where-Object { $_.isMatch() } | ForEach-Object {  Format-PII -File $File  -Match $_; break lines }
                }
            }

        }
        catch{
            if($_.Exception.GetType() -eq [System.UnauthorizedAccessException] ){
                exit
            }

            Write-Log -Message "$_ Attempting Alternate Method..."  -Severity 2 -File $File -FileType "XLSX"
            $XLSXZIP = $File -Replace "xlsx","zip"
            $XLSXUNZIP = $File -Replace ".xlsx",""
                
                try{
                    Copy-Item -Path $File -Destination $XLSXZIP
                    Expand-Archive -Path $XLSXZIP -DestinationPath $XLSXUNZIP
                    
                    $EXCELDOC = Get-ChildItem "$XLSXUNZIP\xl\sharedstrings.xml"
                    [xml]$XmlDocument = Get-Content $EXCELDOC
                    Remove-Item $XLSXZIP,$XLSXUNZIP -Force -Recurse
                    
                    $EXCELTEXT = ($XmlDocument.GetElementsByTagName("t") | Select-Object @{n='Value';e={$_.'#text'}}).value
                    
                    # compile all extracted text
                    $EXCELTEXT | Get-PIIString -FileType "XLSX" | Where-Object { $_.isMatch() } | ForEach-Object {  Format-PII -File $File  -Match $_; break lines }

                }catch{
                    Write-Log -Message "$_ Alternate Method Failed"  -Severity 2 -File $File -FileType "XLSX"
                    Remove-Item $XLSXZIP,$XLSXUNZIP -Force -Recurse -ErrorAction SilentlyContinue
                } 
        }
        finally{
            if($null -ne $spreadsheetDocument)
            {
                $spreadsheetDocument.Close()
            }
        }
    }
}
Export-ModuleMember  -Function Find-XLSX
function Find-DOCX {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [String]$File
    )
    
    begin {
        Try { [IO.File]::OpenWrite($File).Close() }
        Catch { Write-Log -Message "Access Denied" -Severity 3 -File $File -FileType "DOCX"; return }
    }
    
    process {
        try{
            $wordPackage = [System.IO.Packaging.Package]::Open($File, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
    
            [DocumentFormat.OpenXml.Packaging.WordprocessingDocument] $wordDocument = 
                [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($wordPackage)
    
                $wordDocument.MainDocumentPart.Document.Body | ForEach-Object{ 
                        $_.InnerText | Get-PIIString -FileType "DOCX" | Where-Object { $_.isMatch() } | ForEach-Object { Format-PII -File $File  -Match $_; break}
                }
    
        }catch{
            if($_.Exception.GetType() -eq [System.UnauthorizedAccessException] ){
                exit
            }
            Write-Log -File $File -Message "$_ OpenXML Failed. Attempting Alternate Method" -Severity 2 -FileType "DOCX"
            if($wordDocument)
            {
                $wordDocument.Close()
            }

            try{

                $doc = new-object -type system.xml.xmldocument
                $Package = [System.IO.Packaging.Package]::Open($File)
                $documentType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                $OfficeDocRel = $Package.GetRelationshipsByType($documentType)
                $documentPart = $Package.GetPart([System.IO.Packaging.PackUriHelper]::ResolvePartUri("/", $OfficeDocRel.TargetUri))
                $doc.load($documentPart.GetStream())
            
                $mgr = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
                $mgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            
                $nodes = $doc.SelectNodes("/descendant::w:t", $mgr)
                
                :nodes foreach($node in $nodes)
                {
                    $node.InnerText | Get-PIIString -FileType "DOCX" | Where-Object { $_.isMatch() } | `
                        ForEach-Object { Format-PII -File $File  -Match $_; break nodes}
                }
                $Package.Close()
            }
            catch{
                Write-Log -File $File -Message "$_ Alternate Failed" -Severity 2 -FileType "DOCX"
            }
            finally{
             
                if($null -ne $Package){
                    $Package.Close()
                }
            }
        }
        finally{
            if($null -ne $Package){
                $Package.Close()
            }
            
        }
       }

}
Export-ModuleMember  -Function Find-DOCX
function Write-Report {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][String]$Data
    )
    
    begin {
        $head = '<meta http-equiv="x-ua-compatible" content="ie=11"/>'
        $sb = @'
        <script type="text/javascript">
      $(document).ready(function() {
        $("body").addClass("container-fluid");
        $("body").css("margin-top", "20px");

        $("table").each(function() {
          // Grab the contents of the first TR element and save them to a variable
          var tHead = $(this)
            .find("tr:first")
            .html();
          // Remove the first COLGROUP element
          $(this)
            .find("colgroup")
            .remove();
          // Remove the first TR element
          $(this)
            .find("tr:first")
            .remove();
          // Add a new THEAD element before the TBODY element, with the contents of the first TR element which we saved earlier.
          $(this)
            .find("tbody")
            .before("<thead>" + tHead + "</thead>");
        });
        // Apply the DataTables jScript to all tables on the page

        $("table thead tr")
          .clone(true)
          .appendTo("table thead");
        $("table thead tr:eq(1) th").each(function(i) {
          var title = $(this).text();
          $(this).html(
            '<input class="form-control" type="text" placeholder="Search ' +
              title +
              '" />'
          );

          $("input", this).on("keyup change", function() {
            if (table.column(i).search() !== this.value) {
              table
                .column(i)
                .search(this.value)
                .draw();
            }
          });
        });

        var groupColumn = 0;
        var table = $("table").DataTable({
          orderCellsTop: true,
          columnDefs: [
            { visible: false, targets: [0, 1] },
            { width: "5%", name: "Confidence", targets: 2 },
            { width: "25%", name: "File Name", targets: 3 },
            { width: "2%", name: "File Type", targets: 4 },
            { width: "5%", name: "Creation Time", targets: 5 },
            { width: "5%", name: "Last Write Time", targets: 6 },
            { width: "2%", name: "Has SSN", targets: 7 },
            { width: "2%", name: "Has Passport", targets: 8 },
            { width: "2%", name: "Unlocked File", targets: 9 },
            { width: "20%", name: "Owner", targets: 10 },
            { width: "5%", name: "DSN", targets: 11 },
            { width: "10%", name: "Email", targets: 12 },
            { width: "5%", name: "Org", targets: 13 },
            { width: "5%", name: "Base", targets: 14 }
          ],
          order: [[groupColumn, "asc"]],
          displayLength: 25,
          rowGroup: {
            dataSrc: function(row) {
              return (
                '<tr class="h2 bg-warning"><td colspan="42"><a href="'+row[0] +'">' + row[0] +
                "</a></td></tr>"
              );
            }
          },
          bAutoWidth: false,
          bSortClasses: false,
          bDeferRender: true,
          sPagingationType: "full_numbers",
          iDisplayLength: 20
        });
      });
    </script>
'@
        $css = @(
    '<link rel="stylesheet" href="css/dataTables.min.css">',
    '<link rel="stylesheet" href="css/bootstrap.min.css" >',
    '<link rel="stylesheet" href="css/rowGroup.dataTables.min.css" >',
    '<link rel="stylesheet" href="css/responsive.dataTables.min.css" >',
    '<div class="card bg-danger">',
    '<div class="card-header text-center">',
    '<h2>PII Results for ({0})</h2></div><div class="card-body">' -f (Get-Date -DisplayHint Date)
)
        $js = @(
    '</div></div>',
    '<script type="text/javascript" src="js/jquery-3.3.1.min.js" ></script>',
    '<script type="text/javascript" src="js/popper.min.js" ></script>',
    '<script type="text/javascript" src="js/bootstrap.min.js"></script>',
    '<script type="text/javascript" src="js/dataTables.min.js"></script>',
    '<script type="text/javascript" src="js/dataTables.rowGroup.min.js"></script>',
    '<script type="text/javascript" src="js/dataTables.responsive.min.js"></script>',
    $sb
)
    }
    
    process {
        $OutputFolder = Get-Folder
        $OutputFolder = "$OutputFolder\PII_Report\"
        $report = "$script:ReportRoot\report.hta"
        
        #Remove Existing report.csv
        Remove-Item $script:ReportRoot\report.csv -Force -ErrorAction SilentlyContinue
        
        #Get all collected reports
        Get-ChildItem "$script:ReportRoot\*rpt.csv" | ForEach-Object {  Import-Csv -Path $_| Export-csv -Path "$script:ReportRoot\report.csv" -Append -NoTypeInformation }
        
        #compile report
        Import-Csv -Path "$script:ReportRoot\report.csv" | `
            ConvertTo-Html -Body $body -PreContent $css -PostContent $js -Title "PII Searcher" -Head $head | `
            Out-File -Encoding utf8 -FilePath $report
        
        #copy report to location
        Copy-Item -Path "$script:ReportRoot\*" -Recurse -Destination (New-Item $OutputFolder -Type Directory) -Force
    }

    end{
        return $report
    }
    
}
Export-ModuleMember -Function Write-Report
function Write-Log{
    
    Param(
        # Error Message
        [Parameter()]
        [String]
        $Message,
        # Error Severity
        [Parameter()]
        [String]
        $Severity = 'Information',
        # File name 
        [Parameter()]
        [String]
        $File,
        # FileType
        [Parameter()]
        [String]
        $FileType
    )

    [PSCustomObject]@{
        Time = (Get-Date -f g)  
        File = $File
        Message = $Message
        Severity = $Severity
    } | Export-Csv -Path "$LogRoot\PII_$($FileType)_ErrorLog.csv" -Append -NoTypeInformation
}
Export-ModuleMember  -Function Write-Log
function Find-OldNetworkFiles
{
    <#
Created by TSgt Nicholas Crenshaw , 86CS/SCOO

This multithreaded script will report all files older than 1 year stored under a root directory.
This script requires PoshRSJob and PSAlphaFS powershell modules.

#>

$NetworkDriveList = @(

    "\\dm-wg-02\ram-rdr2$"
    "\\dm-usafe-01\86CS_R$"
    "\\dm-usafe-02\ram-rdr1$"
    "\\dm-usafe-01\ram-udr1$"
    "\\dm-tn-01\ram-sdr2$"
    "\\dm-wg-02\fram-rdr7$"
    "\\dm-wg-01\ram-rdr6$"
    "\\DM-PUBLIC\ram-86aw1$"
    "\\dm-wg-02\ram-rdr2$"
    "\\dm-usafe-01\ram-rdr5$"
    "\\dm-usafe-01\ram-odr1$"
    "\\dm-usafe-02\ram-rdr4$"
    "\\dm-public\ram-sdr3$"
    "\\dm-wg-02\603ACOMS_R1$"
    "\\dm-wg-01\603AOC_R1$"
    "\\dm-tn-01\admin-a7a1"
    "\\dm-public\ram-sdr1$"
    "\\dm-wg-01\ram-rdr3$"
    )
    
    foreach($root in $NetworkDriveList){
    
    $LastNetworkDrive = [bool]($root | Select-Object-string -SimpleMatch ($NetworkDriveList | Select-Object -Last 1))
    Do{Start-Sleep 5}
    Until (([bool](Get-RSJob -State Running) -eq $false) -or $LastNetworkDrive)
    
        function CheckRootFolder ($root) {
    
        $date = get-date -Format yyyy-MM-dd
    
        $EXT = @{Name="FileType";Expression={($_.name.split(".") | Select-Object -last 1).tolower() }}
        $KB = @{Name="KBytes";Expression={$Length = ($_.Length / 1KB); if(("{0:N0}" -f $Length) -eq 0){"{0:N2}" -f $Length}else{"{0:N0}" -f $Length}  }}
        $MB = @{Name="MBytes";Expression={"{0:N0}" -f (($_.Length / 1MB) | Where-Object {$_ -ge 1}) }}
        $GB = @{Name="GBytes";Expression={"{0:N2}" -f (($_.Length / 1GB)  | Where-Object {$_ -ge 1}) }}
        $Over1YR = @{Name="Over1YR";Expression={($_.LastWriteTime -lt (Get-Date).AddYears(-1)) }}
        $Over3YR = @{Name="Over3YR";Expression={($_.LastWriteTime -lt (Get-Date).AddYears(-3)) }}
        $Over5YR = @{Name="Over5YR";Expression={($_.LastWriteTime -lt (Get-Date).AddYears(-5)) }}
        $Fullname = @{Name="Path";Expression={($_.fullname)}}
    
    
        $report = Get-LongChildItem $root | Select-Object name, $Fullname, LastWriteTime, length  | Where-Object {$_.LastWriteTime -lt (Get-Date).AddYears(-1)} | Select-Object name, $EXT, fullname, lastaccesstime, lastwritetime, creationtime, $Over1YR, $Over3YR, $Over5YR, $KB, $MB, $GB 
    
        $rootfolder = $root -split "\\" | Select-Object -Last 1
        $file = "$script:ScriptDir\reports\OldFileReport - $rootfolder.csv"
        $report | Export-Csv -Path $file -Append -NoTypeInformation
        }
        CheckRootFolder -root $root
    
    
        Get-ChildItem $root -Directory | Select-Object -ExpandProperty fullname | Start-RSJob -Name {"$($_)"} -Throttle 15 -ScriptBlock{
        param($path)
    
        $date = get-date -Format yyyy-MM-dd
    
        $EXT = @{Name="FileType";Expression={($_.name.split(".") | Select-Object -last 1).tolower() }}
        $KB = @{Name="KBytes";Expression={$Length = ($_.Length / 1KB); if(("{0:N0}" -f $Length) -eq 0){"{0:N2}" -f $Length}else{"{0:N0}" -f $Length}  }}
        $MB = @{Name="MBytes";Expression={"{0:N0}" -f (($_.Length / 1MB) | Where-Object {$_ -ge 1}) }}
        $GB = @{Name="GBytes";Expression={"{0:N2}" -f (($_.Length / 1GB)  | Where-Object {$_ -ge 1}) }}
        $Over1YR = @{Name="Over1YR";Expression={($_.LastWriteTime -lt (Get-Date).AddYears(-1)) }}
        $Over3YR = @{Name="Over3YR";Expression={($_.LastWriteTime -lt (Get-Date).AddYears(-3)) }}
        $Over5YR = @{Name="Over5YR";Expression={($_.LastWriteTime -lt (Get-Date).AddYears(-5)) }}
        $Fullname = @{Name="Path";Expression={($_.fullname)}}
    
        
        $report = Get-LongChildItem $PATH -Recurse | Select-Object name, $Fullname, LastWriteTime, length  | Where-Object {$_.LastWriteTime -lt (Get-Date).AddYears(-1)} | Select-Object name, $EXT, fullname, lastaccesstime, lastwritetime, creationtime, $Over1YR, $Over3YR, $Over5YR, $KB, $MB, $GB
    
        $rootfolder = $using:root -split "\\" | Select-Object -Last 1
        $file = "$script:ScriptDir\reports\OldFileReport - $rootfolder.csv"
    
        Do{
            try { [IO.File]::OpenWrite($file).close();$success = $true
                 $report | Export-Csv -Path $file -Append -NoTypeInformation }
            catch {$success = $false;Start-Sleep 1}}
    
        Until ($success -eq $true)
    
        }
    
        $CurrentDrive = $NetworkDriveList.IndexOf($root) + 1
        $TotalDrives = $NetworkDriveList.count
        
        if($LastNetworkDrive){ Write-Host -ForegroundColor Yellow "Currently scanning network drive $CurrentDrive/$TotalDrives ..."
            Do{Start-Sleep 5
               if((Get-RSJob -State Running).count -eq 0){Write-Host -ForegroundColor Green "Script complete!...All network drives have been scanned."}}
            Until((Get-RSJob -State Running).count -eq 0)
            }
        else{Write-Host -ForegroundColor Yellow "Currently scanning network drive $CurrentDrive/$TotalDrives ..." }
    
    
    }
}
Export-ModuleMember  -Function Find-OldNetworkFiles
function Find-UnauthorizedNetworkFiles
{
    <#
Created by TSgt Nicholas Crenshaw , 86CS/SCOO

This multithreaded script will report all unauthorized files stored under a root directory.
This script requires PoshRSJob and PSAlphaFS powershell modules.

#>

$NetworkDriveList = @(
    "\\dm-wg-02\ram-rdr2$"
    "\\dm-usafe-01\86CS_R$"
    "\\dm-usafe-02\ram-rdr1$"
    "\\dm-tn-01\ram-sdr2$"
    "\\dm-wg-02\fram-rdr7$"
    "\\dm-wg-01\ram-rdr6$"
    "\\DM-PUBLIC\ram-86aw1$"
    "\\dm-usafe-01\ram-rdr5$"
    "\\dm-usafe-01\ram-odr1$"
    "\\dm-usafe-02\ram-rdr4$"
    "\\dm-public\ram-sdr3$"
    "\\dm-wg-02\603ACOMS_R1$"
    "\\dm-wg-01\603AOC_R1$"
    "\\dm-tn-01\admin-a7a1"
    "\\dm-public\ram-sdr1$"
    "\\dm-wg-01\ram-rdr3$"
    "\\dm-wg-02\ram-rdr2$"
    # "\\dm-usafe-01\ram-udr1$"
    )
    
    
    foreach($root in $NetworkDriveList){
    
    $LastNetworkDrive = [bool]($root | Select-Object-string -SimpleMatch ($NetworkDriveList | Select-Object -Last 1))
    Do{Start-Sleep 5}
    Until (([bool](Get-RSJob -State Running) -eq $false) -or $LastNetworkDrive)
    
        function CheckRootFolder ($root) {
    
        $UnauthFileTypes = Get-Content "../dest/config/UnauthorizedFileTypes.config"

        $date = get-date -Format yyyy-MM-dd
    
        $EXT = @{Name="FileType";Expression={($_.name.split(".") | Select-Object -last 1).tolower() }}
        $KB = @{Name="KBytes";Expression={$Length = ($_.Length / 1KB); if(("{0:N0}" -f $Length) -eq 0){"{0:N2}" -f $Length}else{"{0:N0}" -f $Length}  }}
        $MB = @{Name="MBytes";Expression={"{0:N0}" -f (($_.Length / 1MB) | Where-Object {$_ -ge 1}) }}
        $GB = @{Name="GBytes";Expression={"{0:N2}" -f (($_.Length / 1GB)  | Where-Object {$_ -ge 1}) }}
        $Over1YR = @{Name="Over1YR";Expression={($_.LastAccessTime -lt (Get-Date).AddYears(-1)) }}
        $Over3YR = @{Name="Over3YR";Expression={($_.LastAccessTime -lt (Get-Date).AddYears(-3)) }}
        $Over5YR = @{Name="Over5YR";Expression={($_.LastAccessTime -lt (Get-Date).AddYears(-5)) }}
        $Fullname = @{Name="Path";Expression={($_.fullname)}}
    
    
        $report = Get-LongChildItem $root -Include $UnauthFileTypes | Select-Object name, $EXT, $Fullname, lastaccesstime, lastwritetime, creationtime, $Over1YR, $Over3YR, $Over5YR, $KB, $MB, $GB 
    
        $rootfolder = $root -split "\\" | Select-Object -Last 1
        $file = "$script:ScriptDir\reports\UnauthFiles - $rootfolder.csv"
        $report | Export-Csv -Path $file -Append -NoTypeInformation
        }
        CheckRootFolder -root $root
    
        
        Get-ChildItem $root -Directory | Select-Object -ExpandProperty fullname | Start-RSJob -Name {"$($_)"} -Throttle 15 -ScriptBlock{
        param($path)
    
        $UnauthFileTypes = Get-Content "$script:ScriptDir\config\UnauthorizedFileTypes.config"

        $date = get-date -Format yyyy-MM-dd
    
        $EXT = @{Name="FileType";Expression={($_.name.split(".") | Select-Object -last 1).tolower() }}
        $KB = @{Name="KBytes";Expression={$Length = ($_.Length / 1KB); if(("{0:N0}" -f $Length) -eq 0){"{0:N2}" -f $Length}else{"{0:N0}" -f $Length}  }}
        $MB = @{Name="MBytes";Expression={"{0:N0}" -f (($_.Length / 1MB) | Where-Object {$_ -ge 1}) }}
        $GB = @{Name="GBytes";Expression={"{0:N2}" -f (($_.Length / 1GB)  | Where-Object {$_ -ge 1}) }}
        $Over1YR = @{Name="Over1YR";Expression={($_.LastAccessTime -lt (Get-Date).AddYears(-1)) }}
        $Over3YR = @{Name="Over3YR";Expression={($_.LastAccessTime -lt (Get-Date).AddYears(-3)) }}
        $Over5YR = @{Name="Over5YR";Expression={($_.LastAccessTime -lt (Get-Date).AddYears(-5)) }}
        $Fullname = @{Name="Path";Expression={($_.fullname)}}
    
    
    
        $report = Get-LongChildItem $PATH -Recurse -Include $UnauthFileTypes | Select-Object name, $EXT, $Fullname, lastaccesstime, lastwritetime, creationtime, $Over1YR, $Over3YR, $Over5YR, $KB, $MB, $GB 
    
        $rootfolder = $using:root -split "\\" | Select-Object -Last 1
        $file = "$script:ScriptDir\reports\UnauthFiles - $rootfolder.csv"
    
        Do{
            try { [IO.File]::OpenWrite($file).close();$success = $true
                 $report | Export-Csv -Path $file -Append -NoTypeInformation }
            catch {$success = $false;Start-Sleep 1}}
    
        Until ($success -eq $true)
    
        }
    
        $CurrentDrive = $NetworkDriveList.IndexOf($root) + 1
        $TotalDrives = $NetworkDriveList.count
        
        if($LastNetworkDrive){ Write-Host -ForegroundColor Yellow "Currently scanning network drive $CurrentDrive/$TotalDrives ..."
            Do{Start-Sleep 5
               if((Get-RSJob -State Running).count -eq 0){Write-Host -ForegroundColor Green "Script complete!...All network drives have been scanned."}}
            Until((Get-RSJob -State Running).count -eq 0)
            }
        else{Write-Host -ForegroundColor Yellow "Currently scanning network drive $CurrentDrive/$TotalDrives ..." }
    }
}
Export-ModuleMember -Function Find-UnauthorizedNetworkFiles
function Move-UnauthorizedFiles
{

}
Export-ModuleMember -Function Move-UnauthorizedFiles
function Find-Files{
    Param(
        [Parameter(Mandatory=$true)][String[]]$FileTypes
    )
    begin {

        Remove-Item "$ConfigRoot\FileList.txt" -Force -ErrorAction SilentlyContinue

        
    }
    process {
        do{
            Clear-Host
            Write-host "Warning: This process can take a long time. Press 'Enter' to continue 'q' to exit: " -NoNewLine -BackgroundColor Yellow -ForegroundColor Red
            $cont = Read-Host 
            
            if($cont -ne "q")
            {
                    $TargetFolder = Read-Host "`nPlease Enter root folder to scan"
                    if(-not (Test-Path $TargetFolder))
                    {
                        Write-Host "Folder does not exist. Please enter a valid folder"
                        pause
                        $cont=$null
                        continue
                    }
                    else{
                        break
                    }
            }
        }while(-not($cont -eq 'q'))
        filter dirFilter
        {
            switch -Regex ($_)
            {
                'MILCON Projects'{ break }
                'Itemized bills'{ break }
                'Weekly Activity Reports'{ break }
                'WAR Archive'{ break }
                '02. HBSS'{ break }
                '99\.Archive'{ break }
                '99\.Archived'{ break }
                default{$_.TrimStart();break}
            }
        }

        $sb = {
            Param(
                [Parameter(ValueFromPipeline=$true)][String]$dir
            )

            $types = $($FileTypes | ForEach-Object { "*.$_" }) -join " "
            filter fileFilter{
                switch -Regex ($_)
                {
                    'AFI[0-9]{2}[ -]'{ break }
                    'SAV Reports'{ break }
                    'Telekom Mobil'{ break }
                    'Procedures'{ break }
                    'Template'{ break }
                    'Vodafone'{ break }
                    'AFI [0-9]{2}[ -]'{ break }
                    'AFI_[0-9]{2}[ -]'{ break }
                    'AFMAN[0-9]{2}[ -]'{ break }
                    'AFMAN [0-9]{2}[ -]'{ break }
                    'DODI[0-9]{2}[ -]'{ break }
                    'DODI [0-9]{2}[ -]'{ break }
                    'DOD[0-9]{2}[ -]'{ break }
                    'DOD [0-9]{2}[ -]'{ break }
                    'USAFEI [0-9]{2}[ -]'{ break }
                    'USAFEI[0-9]{2}[ -]'{ break }
                    'USAFE[0-9]{2}[ -]'{ break }
                    'USAFE [0-9]{2}[ -]'{ break }
                    'CFETP'{ break }
                    'AFJQS'{break }
                    default{ $_.TrimStart() }
                }
            }
            
            if($dir -ne $Using:TargetFolder){
                & powershell.exe /c "robocopy.exe $dir NULL $types /L /S /NDL /NJH /NJS /NC /NS /NP /R:0 /W:0 /XJ /MT:10" | fileFilter
            }else{
                & powershell.exe /c "robocopy.exe $dir NULL $types /L /NDL /NJH /NJS /NC /NS /NP /R:0 /W:0 /XJ /MT:10" | fileFilter
            }
        } 
        
        $directories = & powershell.exe /c "robocopy.exe $TargetFolder NULL /L /S /LEV:2 /NFL /NJH /NJS /NC /NS /NP /R:0 /W:0 /XJ" | dirFilter

        $directories | Where-Object { $_ -ne "" } | Start-RSJob -ScriptBlock $sb | Out-Null

        Write-Host "[" -NoNewline -ForegroundColor Yellow -BackgroundColor Green
        
        $sw = New-Object System.IO.StreamWriter "$ConfigRoot\FileList.txt"

        while (Get-RSJob){
            Start-Sleep -Milliseconds 100
            Write-Host "0" -NoNewline -ForegroundColor Yellow -BackgroundColor Green
            $Files = Get-RSJob | Where-Object { $_.State -in 'Completed','Failed','Stopped','Suspended','Disconnected'  }
            $Files | Receive-RSJob | Where-Object{ $_ -ne ""}
            $Files | Remove-RSJob
        }
        
        $sw.Close()
        Write-Host "]" -ForegroundColor Yellow -BackgroundColor Green
        Write-Host " Done!" -ForegroundColor Green
    }
}
Export-ModuleMember -Function Find-Files
 function Get-Folder
 {
     $PSDrives = Get-PSDrive | Select-Object name,displayroot | Where-Object {$_.DisplayRoot -ne $null}
     $Topmost = New-Object System.Windows.Forms.Form
     $Topmost.TopMost = $True
     $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
     $foldername.Description = "Select a root folder to scan"
     $foldername.rootfolder = "MyComputer"
 
     if($foldername.ShowDialog($Topmost) -eq "OK")
     {
         $folder += $foldername.SelectedPath
     }
 
     if($PSDrives.name -contains $folder[0]){
     $displayroot = $PSDrives | Where-Object {$_.name -match $folder[0]} | Select-Object -ExpandProperty displayroot
     $folder =  $displayroot + $folder.Substring(2)}
     
     return $folder
 }
 Export-ModuleMember -Function Get-Folder
 function Get-PIIString {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [AllowEmptyString()][String] $TextData,
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] 
        [AllowEmptyString()][String] $FileType
    )
    
    begin {
        if($null -eq $TextData -or "" -eq $TextData){return}
        }
    
    process {
        $scriptBlock_1 = {
            Param(
                [String]$Text
            )
            $scriptBlock_2 = {
                Param(
                    [psobject]$matches
                )
    
                $m = [PIIMatch]::new()
                $m.FileType = $FileType
            
                switch ($matches.Keys | Select-Object -First 1)
                { 
                    
                    "SSN"  {
                        $m.HasSSN = $true
                        $m.confidence = "High"
                        break
                    }
                    "Passport"  {
                        $m.HasPassport = $true
                        $m.confidence = "High"
                        break
                    }
                    default  {
                        $m.confidence = "Medium"
                        break
                    }
                }
                return $m
            }
          
            switch -Regex ($Text)
            {
                "\b(?<SSN>(SSAN|SSN|Social Security):?\s?(?!000)(?!666)(?!9)[0-9]{3}[ -](?!00)[0-9]{2}[ -](?!0000)[0-9]{4})\b"{& $scriptBlock_2 $matches ;break}
                "\b(?<SSN>(SSAN|SSN|Social Security):?\s?(?!(000|666|9))\d{3}[ -](?!00)\d{2}[ -](?!0000)\d{4})\b"{& $scriptBlock_2 $matches ;break}
                "\b(?<Passport>Passport:?\s?[a-zA-Z]{2}[0-9]{7})\b"{& $scriptBlock_2 $matches ;break}
                default { & $scriptBlock_2 $Text; break}
            }
        }
  
        switch -Regex ($TextData)
        {
            "\b(?<Text>Social\sSecurity)\b"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>Social\sSecurity\sNumber:?)\b"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>AROWS\sTRACKING\sNUMBER)\b"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>AF\sFORM\s899)\b"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>SSAN)\b"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>ssan)\b"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>SSN:?)"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>ssn:?)"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>Passport:?)"{& $scriptBlock_1 $TextData ;break}
            "\b(?<Text>passport:?)"{& $scriptBlock_1 $TextData ;break}
        }
    }

}
Export-ModuleMember -Function Get-PIIString
function Find-PIIOwner {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String]$File
    )
    
    begin {
        filter sidFilter
        {
            #Get the EDIP from the Owner field
            if($_ -match "\d{10}[a-zA-Z]"){($Matches | Select-Object -First 1 ).Values }
        }
    }
    
    process {
        try{
            $RestrictedSecGroups = "Everyone","All","Authenticated","Users"
            $securitygroups = (Get-Acl $File | ForEach-Object {$_.Access} | Select-Object -ExpandProperty identityreference).value
            $UnlockedGroups = $RestrictedSecGroups | ForEach-Object {$securitygroups -match $_}
            $UnlockedFile = [bool]($UnlockedGroups)
            
            $O = Get-Acl -Path $File | Select-Object -Property Owner | sidFilter 
            if($null -ne $O){
                $a = [adsisearcher]""
                $a.Filter = "samAccountName={0}" -f $O
                $Owner = $a.FindOne() | Select-Object -ExpandProperty Properties
            }
        }
        catch
        {
            Write-Log -Severity 2 -File $File -Message $_ -FileType "General"
        }
    }
    
    end {
        $retVal = [PIIOwnerInfo]::new()
        $retVal.DisplayName    = $Owner.displayname
        $retVal.OfficePhone    = $Owner.telephonenumber
        $retVal.EmailAddress   = $Owner.mail
        $retVal.Org            = $Owner.o
        $retVal.Base           = $Owner.l
        $retVal.UnlockedFile = $UnlockedFile
        $retVal.SecurityGroups = $securitygroups
        $retVal.UnlockedGroups = $UnlockedGroups
    
        return $retVal
    }
}
Export-ModuleMember -Function Find-PIIOwner

function Format-PII {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [PIIMatch]$match,
        [Parameter(Mandatory=$true)] [string]$File
    )

    if( $match.isMatch()){

        [PIIOwnerInfo]$Owner = Find-PIIOwner -File $File
        [PII]$retVal = [PII]::new()
        
        $retVal.RootFolder     = $File | Split-Path -Parent
        $retVal.FileType       = $match.Filetype
        $retVal.FileName       = $File | Split-Path -Leaf
        $retVal.FullPath       = $File
        $retVal.confidence     = $match.confidence
        $retVal.CreationTime   = (Get-Item $File).CreationTime 
        $retVal.LastAccessTime = (Get-Item $File).LastWriteTime
        $retVal.HasSSN         = [bool]$match.HasSSN
        $retVal.HasPassport    = [bool]$match.HasPassport
        $retVal.UnlockedFile   = $Owner.UnlockedFile
        $retVal.DisplayName    = $Owner.DisplayName
        $retVal.OfficePhone    = $Owner.OfficePhone
        $retVal.EmailAddress   = $Owner.EmailAddress
        $retVal.Org            = $Owner.Org        
        $retVal.Base           = $Owner.Base       
        
        Write-Log -Severity 2 -File $File -Message $Owner.mail -FileType "General"
        return $retval
    }
}
Export-ModuleMember -Function Format-PII
function Get-TessTextFromImage {
    Param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="ImageObject")][System.Drawing.Image]$Image,
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="FilePath")][String]$Path
    )

    $tesseract = New-Object Tesseract.TesseractEngine((Get-Item "$script:ScriptDir\bin\tessdata").FullName, "eng", [Tesseract.EngineMode]::Default, $null)

	Get-Process {
		#load image if path is a param
		If ($PsCmdlet.ParameterSetName -eq "FilePath") { $Image = New-Object System.Drawing.Bitmap((Get-Item $path).Fullname) } 

		#perform OCR on image
		$pix = [Tesseract.PixConverter]::ToPix($image)
		$page = $tesseract.Process($pix)
	
		#build return object
		$ret = New-Object PSObject -Property @{"Text"= $page.GetText();
										   "Confidence"= $page.GetMeanConfidence()}

		#clean up references
		$page.Dispose()
		If ($PsCmdlet.ParameterSetName -eq "FilePath") { $image.Dispose() } 
		return $ret
	}
}
Export-ModuleMember -Function Get-TessTextFromImage
function Find-PII{
    Param(
        # File Type Filter
        [Parameter(Mandatory=$true)][String[]]$FileTypes
    )

    
    Write-Host "Searching for PII...."

    $sb = {
        Param(
            [Parameter()][String]$File
        )
        #Colect the file listing from the configuration directory

        switch -Regex ($File)
        {
            "DOCX$"  { Find-DOCX -File $File     }
            "XLSX$"  { Find-XLSX -File $File     }
            "PDF$"   { Find-PDF -File $File      }
            "CSV$"   { Find-FlatFile -File $File }
        }
    }

    $fileList = "$ConfigRoot\FileList.txt"

    [System.IO.File]::ReadLines($fileList) | Start-RSjob -Name $type -ScriptBlock $sb -ModulesToImport @("PIIFinding") -ArgumentList $type -Throttle 10 | Out-Null
    
    Write-Host "`r`rWarning: This is  a long running process." -BackgroundColor Red -ForegroundColor White
    Write-Host "Discovering PII: " -NoNewline -ForegroundColor Red -BackgroundColor Green

    $sw = New-Object System.IO.StreamWriter "$ReportRoot\report.txt"

    while(Get-RSJob)
    {
        Start-Sleep -Milliseconds 100
        Write-Host "." -NoNewline -ForegroundColor Green
        $Files = Get-RSJob | Where-Object { $_.State -in 'Completed','Disconnected','Failed','Stopped','Stopping' } | Receive-RSJob | ForEach-Object{ $sw.WriteLine($_)}
        $Files | Remove-RSJob
    }

    Write-Host "Done!" -ForegroundColor Green
 }
 Export-ModuleMember -Function Find-PII

function Show-Menu
{

    $menu = @'
                                              \     /
                                      x________\(O)/________x
                                          o o  O(.)O  o  o
                          *********************************************
                          *********** Drive Cleanup Scripts ***********
                          *********************************************
                          **                                         **
                          *   1. Discover Files to be scanned         *
                          *                                           *
                          *   2. Scan For PII                         *
                          *                                           *
                          *   3. Create PII Report                    *
                          *                                           *
                          *   4. Report Old Network Files             *
                          *                                           *
                          *   5. Report Unauthorized Network Files    *
                          *                                           *
                          *   6. Move Unauthorized Network Files      *
                          *                                           *
                          *   Q. Press 'Q' to quit                    *
                          *                                           *
                          *********************************************
'@


    do
    {
        Clear-Host
        Write-Host $menu -ForegroundColor Yellow

        $input = Read-Host "`n                           Please make a selection"
        switch ($input)
        {
            '1' {
                    Clear-Host               
                    Find-Files -FileTypes @("XLSX","DOCX","PDF","CSV")
            }'2'{
                    Clear-Host
                    Find-PII -FileTypes @("XLSX","DOCX","PDF","CSV")        
            }'3'{
                    Clear-Host
                    Write-Report -Data  "$ReportRoot" | Invoke-Item
            }'4' {
                    Clear-Host
                    Find-OldNetworkFiles
            }'5'{
                    Clear-Host
                    Find-UnauthorizedNetworkFiles
                }
                '6'{
                    Clear-Host
                    Move-UnauthorizedFiles
                }
            'q'{
                    exit
            }
        }
        Write-Host `n
        pause
    }until($input -eq 'q')

}
Export-ModuleMember -Function Show-Menu


if ("PII" -as [type]) {
	class PII
    {
        [String] $RootFolder
        [String] $FullPath
        [String] $confidence
        [String] $FileName            
        [String] $FileType        
        [String] $CreationTime     
        [String] $LastAccessTime
        [String] $HasSSN        
        [String] $HasPassport   
        [String] $UnlockedFile  
        [String] $DisplayName   
        [String] $OfficePhone   
        [String] $EmailAddress  
        [String] $Org  
        [String] $Base  
    }
}

if ("PIIOwnerInfo" -as [type]) {
    class PIIOwnerInfo
    {
        [String]$DisplayName
        [String]$OfficePhone
        [String]$EmailAddress
        [String]$Org        
        [String]$Base       
        [bool]$UnlockedFile
        [String[]]$SecurityGroups
        [String[]]$UnlockedGroups
    }
}

if ("PIIMatch" -as [type]) {
    class PIIMatch
    {
        PIIMatch(){
            $this.HasSSN = $false
            $this.HasPassport = $false
            $this.confidence = "Low"
            $this.otherValue = $null
        }
        [bool]$HasSSN
        [bool]$HasPassport
        [String]$confidence
        [String]$FileType
        [String]$otherValue
        [bool]isMatch(){
            return (($this.HasSSN -eq $true) -or ($this.HasPassport -eq $true))
        }
    }
}