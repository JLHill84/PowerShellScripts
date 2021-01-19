# This script will not log you in to AZ
# So be sure to login before executing
$excel = New-Object -ComObject excel.application
$excel.visible = $True
$workbook = $excel.Workbooks.Add()
$workbook.Worksheets.Item(1).Name = "Resources"
$resourceSheet = $workbook.Worksheets.Item(1)
$cell = $resourceSheet.Cells
$resourceGroups = Get-AzResource

$nameCol = 1
$cell.Item(1, $nameCol) = "Name"
$resourceSheet.Columns(1).ColumnWidth = 90

$resourceGroupNameCol = 2
$cell.Item(1, $resourceGroupNameCol) = "Resource Group"
$resourceSheet.Columns(2).ColumnWidth = 50

$resourceTypeCol = 3
$cell.Item(1, $resourceTypeCol) = "Resource Type"
$resourceSheet.Columns(3).ColumnWidth = 45

$locationCol = 4
$cell.Item(1, $locationCol) = "Location"
$resourceSheet.Columns(4).ColumnWidth = 10

$tagsCol = 5
$cell.Item(1, $tagsCol) = "Tags"
$resourceSheet.Columns(5).ColumnWidth = 65

$resourceIdCol = 6
$cell.Item(1, $resourceIdCol) = "Resource Id"
$resourceSheet.Columns(6).ColumnWidth = 210

$rowCount = 3
foreach ($group in $resourceGroups) {
    $cell.Item($rowCount, $nameCol) = $group.Name
    $cell.Item($rowCount, $resourceGroupNameCol) = $group.ResourceGroupName
    $cell.Item($rowCount, $resourceTypeCol) = $group.ResourceType -replace "Microsoft."
    $cell.Item($rowCount, $locationCol) = $group.Location
    $cell.Item($rowCount, $resourceIdCol) = $group.ResourceId -replace "/subscriptions/"
    $tags = $group.Tags
    $tagIndex = 0
    # Tags aren't quite right
    if ($tags.length -gt 0) {
        Write-Host $tags
        foreach ($tag in $tags) {
            $cell.Item($rowCount, $tagsCol) = "$($tag.Keys[$tagIndex]): $($tag.Values[$tagIndex]);"
            $tagIndex++
        }
    }
    $rowCount++
}
$workingDir = Get-Location
$path = Join-Path -Path $workingDir -ChildPath "Resources.xlsx"
$excel.DisplayAlerts = $false
$workbook.SaveAs($path)
$excel.Close
