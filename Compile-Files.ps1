function Convert-WordToHTML($docSrc, $htmlOutputPath ){
  $wdTypes = Add-Type -AssemblyName 'Microsoft.Office.Interop.Word' -Passthru
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.Word.WdSaveFormat')
  $srcFiles = Get-ChildItem $docSrc -filter "*.docx" -Recurse
  $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");
  $wordApp = new-object -comobject word.application
  $wordApp.Visible = $false

  ForEach ($doc in $srcFiles) {
    $openDoc = $wordApp.documents.open($doc.FullName);

    $newName = $doc.BaseName -replace ".docx", ""
    $newName = $newName -replace "'", ""
    Write-Host "Converting to html :" $doc.BaseName, "`ninto", "$htmlOutputPath\$newName.html"

    $openDoc.saveas([ref]"$htmlOutputPath\$newName.html", [ref]$saveFormat);
    $openDoc.close();
    $doc = $null
    }
    $wordApp.quit();           
}
 
function Get-WordContents($docSrc){
  $srcFiles = Get-ChildItem $docSrc -filter "*.docx" -Recurse
  $wordApp = New-Object -ComObject Word.application
  ForEach ($doc in $srcFiles) {
    Write-Host "Opening ", $doc.FullName
    $document = $wordApp.Documents.Open($doc.FullName)
    $document.Visible = $false
    $documentContents | Add-Content "C:\Temp\All_Content.txt"
    $document.Close($false)           
  }
  $wordApp.Quit()
}

$startPath = Read-Host "Where would you like me to start - absolute path only please? " 
pushd $startPath

# Get the folders
Write-Host "Looking for folders"
$folders = @()
$folders_with_files = @{}
Get-ChildItem * -Recurse -Directory -Depth 2 | ForEach-Object {
  if(($_.BaseName.ToLower() -match "how" -or $_.BaseName.ToLower() -match "guide") -and $_.FullName -notcontains "Department_Wiki" -and $_.BaseName -notcontains "Archive" -and $_.BaseName -notcontains "Draft" -and $_.BaseName -notcontains "draft" -and $_.BaseName -notcontains "archive"){
   Write-Host $($_.FullName | Resolve-Path -Relative)
   $folders += $($_.FullName | Resolve-Path -Relative)
   $folders_with_files[$($_.FullName | Resolve-Path -Relative)] = $_.FullName
 }
} # Only get the folders, because the function will get the files in the folder!

# Move to the wiki folder
Write-Host "Moving to the wikipedia folder"
$wikiPath = # Your path here
cd $wikiPath
 
# Recreate the folder structure
Write-Host "Making folders and files"
foreach($folder in $folders){     
  Write-Host "I would like to made a folder in " $pwd + $folder
  mkdir $folder
  Write-Host "Making files for", $folder
  $origionalLocation = $folders_with_files[$folder]
  Convert-WordToHTML $origionalLocation $($wikiPath + $folder) # Function will go and get the files
  Write-Host "Getting contents for", $folder
  Get-WordContents $origionalLocation
}

# Turn it into the website front end
$templateLocation = ".\Department_Wiki\src\"

cp ..\src\style.css .
cp ..\src\script.js .

$all_systems = Get-ChildItem -Directory
$html ="<div id='systems' class=''>"

foreach($system in $all_systems){
  $systemID = $system -replace " ", "_"
  $html += "<h2>$($system)</h2>`n<h3>How To:</h3>`n<ol><div id='$($systemID)' class='system-links'>"
  cd ".\$($system)"
  $files_in_folder = Get-ChildItem -Recurse -File -Filter "*.html" | ForEach-Object{
    $file = $_.BaseName
    $filePath = $($_.FullName | Resolve-Path -Relative)
    if($file -eq"header" -or $file -eq "plchdr"){
      $ignore = $true
    }
    if($filePath -contains "'"){
      $filePath = $filePath -replace "'", "\'"
    }
    $LinkText = '<li><a href="' + "$($system)/$($filePath)" + '"' + ">$($file -replace 'html', '')</a></li>`n"
    if($ignore -eq $false){
      $html += $LinkText
    }
    $ignore = $false
  }

  $html += "</div></ol>"
  cd $wikiPath
}

$html += "</div>"
$htmlContents = Get-Content "$templateLocation\template.html"
$newHTMLConents = $htmlContents -replace "<div id='functions' class=''></div>", $html
$newHTMLConents | Set-Content "index.html"
popd
