#Instruction of use
[System.Environment]::NewLine
Write-Host 'To use this tool, create a CSV file following the example below:'
$example = @(
    @{CollectionName = 'Collection Name 1'; WeekOrder = 'Second'; DayOfWeek = 'Thursday' },
    @{CollectionName = 'Collection Name 2'; WeekOrder = 'Third'; DayOfWeek = 'Wednesday' }) | 
ForEach-Object { New-Object object | Add-Member -NotePropertyMembers $_ -PassThru }

$example | Select-Object CollectionName, WeekOrder, DayofWeek | Format-Table

Read-Host -Prompt 'Press any key to continue or CTRL+C to quit'

#Open file browser to select input CSV file of build info
$sourceFileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$sourceFileBrowser.Filter = 'csv (*.csv)| *.csv'
$sourceFileBrowser.Title = 'Select the input csv file'
$sourceFileBrowser.InitialDirectory = $logDir

if ($sourceFileBrowser.ShowDialog() -eq 'OK') {
    $collectionCsv = Import-Csv ('FileSystem::' + "$($sourceFileBrowser.FileName)")
}
else {
    Write-Log -Message 'CSV selection cancelled by user. Exiting...' -Level ERROR
    break
}

#Validate data in CSV file
Write-Log -Message 'Validating the data entered in the CSV file. Please wait...' -Level INFO
$validationFail = 0
foreach ($col in $collectionCsv) {
    if (-not (($col.CollectionName -and $col.WeekOrder) -or (!$col.CollectionName -and !$col.WeekOrder))) {
        Write-Log -Message 'There is a blank field in the csv file!' -Level ERROR
        $validationFail += 1
    }

    $allowedWeeks = 'First', 'Second', 'Third', 'Fourth', 'Last'
    if ($allowedWeeks -notcontains $col.WeekOrder) {
        Write-Log -Message "$($col.CollectionName) has an invalid week order!" -Level ERROR
        $validationFail += 1
    }

    $allowedDays = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'
    if ($allowedDays -notcontains $col.DayOfWeek) {
        Write-Log -Message "$($col.CollectionName) has an invalid day of the week!" -Level ERROR
        $validationFail += 1
    }
}

if ($validationFail -eq 0) {
    Write-Log -Message 'Csv validation has passed.' -Level SUCCESS
}
else {
    Write-Log -Message 'Csv validation has failed. Please check the errors above. Exiting...' -Level ERROR
    break
}

#Update maintenance window
foreach ($item in $collectionCsv) {
    try {
        Write-Host "$($(Get-Date).toString('HH:mm:ss')) INFO Updating " -NoNewline; Write-Host -ForegroundColor Yellow "$($item.CollectionName)" -NoNewline; Write-Host ' to ' -NoNewline; Write-Host -BackgroundColor Red "the $($item.WeekOrder) $($item.DayOfWeek)"
        $cm = Get-CMCollection -Name $item.CollectionName
        $maintenanceWindow = Get-CMMaintenanceWindow -CollectionId $cm.CollectionID
        $sched = New-CMSchedule -Start $maintenanceWindow.StartTime -End $maintenanceWindow.StartTime.AddMinutes($maintenanceWindow.Duration) -DayOfWeek $item.DayOfWeek -WeekOrder $item.WeekOrder
        
        Set-CMMaintenanceWindow -Name $maintenanceWindow.Name -CollectionId $cm.CollectionID -Schedule $sched
    }
    catch {
        Write-Log -Message "Something went wrong with updating $($item.CollectionName)" -Level ERROR
    }
}

#Validate the changes
[System.Environment]::NewLine
Write-Log -Message 'Validating the changes...' -Level INFO
$finalValidation = 0
foreach ($obj in $collectionCsv) {
    $cm = Get-CMCollection -Name $obj.CollectionName
    $maintenanceWindow = Get-CMMaintenanceWindow -CollectionId $cm.CollectionID
    if (($maintenanceWindow.Description.ToLower().Contains($obj.WeekOrder.toLower())) -and ($maintenanceWindow.Description.ToLower().Contains($obj.DayOfWeek.toLower()))) {
        Write-Log -Message "| $($obj.WeekOrder) $($obj.DayOfWeek) | $($cm.Name)" -Level SUCCESS
    }
    else {
        Write-Log -Message "| $($cm.Name)" -Level ERROR
        $finalValidation += 1
    }
}

if ($finalValidation -eq 0) {
    Write-Log -Message 'Finished! Everything was a success. Exiting.' -Level SUCCESS
}
else {
    Write-Log -Message 'There was a failure in the validation. Please check. Exiting.' -Level ERROR
}