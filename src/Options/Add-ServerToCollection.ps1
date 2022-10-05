Write-Log -Message 'Loading the Add Server to Collection Tool... Please wait.' -Level INFO

# Get all collections and servers from MECM to load into form
$masterScheduleCollection = (Get-WmiObject -ComputerName $SITESERVER -Namespace root/SMS/site_$SITECODE -Class SMS_Collection -Filter "ObjectPath = '$SCHEDULECOLLECTION'").Name
$allWindowsServers = Get-CMCollectionMember -CollectionName $ALLSERVERSCOLLECTION | Select-Object Name

$form1 = New-Object System.Windows.Forms.Form
$form1.ClientSize = New-Object System.Drawing.Size(550, 300)
$form1.TopMost = $true
$form1.MaximizeBox = $false
$form1.FormBorderStyle = 'Fixed3D'
$form1.StartPosition = 'CenterScreen'
$form1.Text = 'Add Server to Collection'

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
$masterScheduleCollection | ForEach-Object { [void] $comboBox1.Items.Add($_) }
$comboBox1.SelectedIndex = 0
$form1.Controls.Add($comboBox1)

$addButton1 = New-Object System.Windows.Forms.Button
$addButton1.Location = New-Object System.Drawing.Point(20, 135)
$addButton1.Size = New-Object System.Drawing.Size(75, 23)
$addButton1.Text = 'Add'
$form1.AcceptButton = $addButton1
$form1.Controls.Add($addButton1)

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
$outputLabel1.Text = 'Welcome to Auto MECM: Add Server to Collection'
$form1.Controls.Add($outputLabel1)

$addButton1.add_Click({
        if ($textBox1.TextLength -eq 0) {
            $outputLabel1.ForeColor = [Drawing.Color]::Red
            $outputLabel1.Text = 'ERROR: Server name cannot be left blank'
            Write-Log -Message 'ERROR: Auto-MECM Add Server to Collection - textbox server entry was left blank' -Level DEBUG
            return
        }

        try {
            $outputLabel1.ForeColor = [Drawing.Color]::Black
            $outputLabel1.Text = "Adding $($textBox1.Text) to $($comboBox1.SelectedItem). Please wait..."

            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $comboBox1.SelectedItem -ResourceId (Get-CMDevice -Name $textBox1.Text).ResourceID
        
            $outputLabel1.ForeColor = [Drawing.Color]::Green
            $outputLabel1.Text = "Successfully added $($textBox1.Text) to $($comboBox1.SelectedItem)!"
            $textBox1.Clear()
            Write-Log -Message "Auto-MECM: Add Server to Collection - Successfully added server $($textBox1.Text) to collection $($comboBox1.SelectedItem)" -Level DEBUG
        }
        catch {
            $outputLabel1.ForeColor = [Drawing.Color]::Red
            $outputLabel1.Text = 'ERROR: Please check the server exists'
            Write-Log -Message "ERROR: Auto-MECM Add Server to Collection - $($textBox1.Text) server cannot be found in MECM" -Level DEBUG
            return
        }
    })

$formResponse = $form1.ShowDialog()

if ($formResponse -ne [System.Windows.Forms.DialogResult]::OK) {
    break
}