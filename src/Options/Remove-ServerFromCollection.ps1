Write-Log -Message 'Loading the Remove Server from Collection Tool... Please wait.' -Level INFO

# Get all servers from MECM to load into form
$allWindowsServers = Get-CMCollectionMember -CollectionName $ALLSERVERSCOLLECTION | Select-Object Name

$form1 = New-Object System.Windows.Forms.Form
$form1.ClientSize = New-Object System.Drawing.Size(550, 300)
$form1.TopMost = $true
$form1.MaximizeBox = $false
$form1.FormBorderStyle = 'Fixed3D'
$form1.StartPosition = 'CenterScreen'
$form1.Text = 'Remove Server from Collection'

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(20, 40)
$textBox1.Size = New-Object System.Drawing.Size(180, 20)
$textBox1.MaxLength = 25
$textBox1.AutoCompleteSource = 'CustomSource' 
$textBox1.AutoCompleteMode = 'SuggestAppend' 
$textBox1.AutoCompleteCustomSource.AddRange($allWindowsServers.Name)
$form1.Controls.Add($textBox1)

$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = New-Object System.Drawing.Point(20, 95)
$comboBox1.Size = New-Object System.Drawing.Size(510, 25)
$comboBox1.DropDownStyle = 'DropDownList'
$comboBox1.IntegralHeight = $false
$comboBox1.MaxDropDownItems = 12
$comboBox1.Sorted = $true
$comboBox1.Enabled = $false
$form1.Controls.Add($comboBox1)

$checkButton1 = New-Object System.Windows.Forms.Button
$checkButton1.Location = New-Object System.Drawing.Point(210, 38)
$checkButton1.Size = New-Object System.Drawing.Size(75, 23)
$checkButton1.Text = 'Check'
$form1.AcceptButton = $checkButton1
$form1.Controls.Add($checkButton1)

$removeButton1 = New-Object System.Windows.Forms.Button
$removeButton1.Location = New-Object System.Drawing.Point(20, 135)
$removeButton1.Size = New-Object System.Drawing.Size(75, 23)
$removeButton1.Text = 'Remove'
$form1.Controls.Add($removeButton1)

$exitButton1 = New-Object System.Windows.Forms.Button
$exitButton1.Location = New-Object System.Drawing.Point(233, 245)
$exitButton1.Size = New-Object System.Drawing.Size(75, 23)
$exitButton1.Text = 'Exit'
$exitButton1.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form1.Controls.Add($exitButton1)

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(20, 20)
$label1.Size = New-Object System.Drawing.Size(150, 23)
$label1.Text = 'Enter a server name:'
$form1.Controls.Add($label1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(20, 75)
$label2.Size = New-Object System.Drawing.Size(150, 23)
$label2.Text = 'Choose a collection:'
$form1.Controls.Add($label2)

$outputLabel1 = New-Object System.Windows.Forms.Label
$outputLabel1.Location = New-Object System.Drawing.Point(20, 175)
$outputLabel1.Size = New-Object System.Drawing.Size(510, 40)
$outputLabel1.MaximumSize = New-Object System.Drawing.Size(510, 40)
$outputLabel1.AutoSize = $true
$outputLabel1.Text = 'Welcome to Auto MECM: Remove Server from Collection'
$form1.Controls.Add($outputLabel1)

$checkButton1.add_Click({
        if ($textBox1.TextLength -eq 0) {
            $outputLabel1.ForeColor = [Drawing.Color]::Red
            $outputLabel1.Text = 'ERROR: Server name cannot be left blank'
            Write-Log -Message 'ERROR: Auto-MECM Remove Server from Collection - textbox server entry was left blank' -Level DEBUG
            return
        }

        $comboBox1.Items.Clear()
        $outputLabel1.ForeColor = [Drawing.Color]::Black
        $outputLabel1.Text = "Checking which collections $($textBox1.Text) is in. Please wait..."

        $serverCollection = (Get-WmiObject -ComputerName $SITESERVER -Namespace root/SMS/site_$SITECODE -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$($textBox1.Text)' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID") | Where-Object { $_.ObjectPath -eq $SCHEDULECOLLECTION }
        if ($serverCollection) {
            $comboBox1.Enabled = $true
            $outputLabel1.ForeColor = [Drawing.Color]::Black
            $outputLabel1.Text = "Found $($textBox1.Text) in $(($serverCollection.Name | Measure-Object).Count) collections"
    
            $serverCollection.Name | ForEach-Object { [void] $comboBox1.Items.Add($_) }
            $comboBox1.SelectedIndex = 0
            Write-Log -Message "Auto-MECM: Remove Server from Collection - Found $($textBox1.Text) in $(($serverCollection.Name | Measure-Object).Count) collections" -Level DEBUG    
        }
        else {
            $outputLabel1.ForeColor = [Drawing.Color]::Red
            $outputLabel1.Text = "ERROR: Could not find $($textBox1.Text) in any collections"
            Write-Log -Message "Auto-MECM: Remove Server from Collection - $($textBox1.Text) server cannot be found in any collections" -Level DEBUG
        }
    })


$removeButton1.add_Click({
        if ($textBox1.TextLength -eq 0) {
            $outputLabel1.ForeColor = [Drawing.Color]::Red
            $outputLabel1.Text = 'ERROR: Server name cannot be left blank'
            Write-Log -Message 'ERROR: Auto-MECM Remove Server from Collection - textbox server entry was left blank' -Level DEBUG
            return
        }

        if (!($comboBox1.Items)) {
            $outputLabel1.ForeColor = [Drawing.Color]::Red
            $outputLabel1.Text = 'ERROR: No collection has been selected'
            Write-Log -Message 'ERROR: Auto-MECM Remove Server from Collection - no collection has been selected' -Level DEBUG
            return
        }

        try {
            $outputLabel1.ForeColor = [Drawing.Color]::Black
            $outputLabel1.Text = "Removing $($textBox1.Text) from $($comboBox1.SelectedItem). Please wait..."

            Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $comboBox1.SelectedItem -ResourceId (Get-CMDevice -Name $textBox1.Text).ResourceID -Force -Confirm:$false
        
            $outputLabel1.ForeColor = [Drawing.Color]::Green
            $outputLabel1.Text = "Successfully removed $($textBox1.Text) from $($comboBox1.SelectedItem)!"

            $comboBox1.Items.Remove($comboBox1.SelectedItem)
            if ($comboBox1.Items) {
                $comboBox1.SelectedIndex = 0
            }
            else {
                $comboBox1.Enabled = $false
            }
            Write-Log -Message "Auto-MECM: Remove Server from Collection - Successfully removed server $($textBox1.Text) from collection $($comboBox1.SelectedItem)" -Level DEBUG
        }
        catch {
            $outputLabel1.ForeColor = [Drawing.Color]::Red
            $outputLabel1.Text = 'ERROR: Please check the server exists'
            Write-Log -Message "ERROR: Auto-MECM Remove Server from Collection - $($textBox1.Text) server cannot be found in MECM or something else went wrong" -Level DEBUG
            return
        }
    })

$formResponse = $form1.ShowDialog()

if ($formResponse -ne [System.Windows.Forms.DialogResult]::OK) {
    break
}