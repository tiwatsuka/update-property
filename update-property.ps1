Clear-Host

if($args.Count -ne 1){
    Write-Host -ForegroundColor Red "Invalid arguments."
    exit 1
}

Write-Host -ForegroundColor Cyan "Reading a property file..."
if(!(Test-Path $args[0])){
    Write-Host -ForegroundColor Red $args[0] "is not found."
    exit 1
}
$customProps = Import-Csv $args[0]
ForEach($prop in $customProps){
    Write-Host -ForegroundColor White " * " $prop.NAME ":" $prop.VALUE
}

Write-Host -ForegroundColor Cyan "Search word files..."
$files = Get-ChildItem -Recurse -Path ./ -include "*.docx" | Select-Object FullName
ForEach($file in $files){
    Write-Host -ForegroundColor White " * " $file.FullName
}

Write-Host -ForegroundColor Cyan "Loading application..."
$application = New-Object -ComObject word.application
$application.Visible = $false

function UpdateProperty($doc, $name, $value){
    $props = $doc.CustomDocumentProperties
    $type = $props.GetType()
    $binding = "System.Reflection.BindingFlags" -as [type]

    Try{
        $propObj = $type.InvokeMember("Item", $binding::GetProperty, $null, $props, $name)
        $type.InvokeMember("Value", $binding::SetProperty, $null, $propObj, $value)
        Write-Host -ForegroundColor White "property" $name "is updated."
    } Catch {
        Write-Host -ForegroundColor White "property" $name "is skipped."
    }
}

ForEach($file in $files){
    Write-Host -ForegroundColor Cyan "Open document..." $file.FullName
    $document = $application.documents.open($file.FullName)

    ForEach($prop in $customProps){
        UpdateProperty $document $($prop.NAME) $($prop.VALUE)
    }

    Write-Host -ForegroundColor Cyan "Updating document field."
    ForEach($section in $document.Sections){
        $document.Fields.Update() | Out-Null

        $headers = $section.Headers
        ForEach($header in $headers){
            $fields = $header.Range.Fields
            ForEach($field in $fields){
                $field.Update() | Out-Null
            }
        }

        $footers = $section.Footers
        ForEach($footer in $footers){
            $fields = $footer.Range.Fields
            ForEach($field in $fields){
                $field.Update() | Out-Null
            }
        }
    }


    Write-Host -ForegroundColor Cyan "Closing document."
    $document.Saved = $false
    $document.save()
    $document.close()
}

$application.quit
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
Write-Host -ForegroundColor Cyan "Done."