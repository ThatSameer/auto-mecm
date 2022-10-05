$form1 = New-Object System.Windows.Forms.Form
$form1.Text = 'Find Servers Collections'
$form1.Size = New-Object System.Drawing.Size(970, 410)
$form1.StartPosition = 'CenterScreen'
$form1.MaximizeBox = $false
$form1.FormBorderStyle = 'Fixed3D'
$form1.Topmost = $true

$findButton1 = New-Object System.Windows.Forms.Button
$findButton1.Location = New-Object System.Drawing.Point(238, 100)
$findButton1.Size = New-Object System.Drawing.Size(75, 23)
$findButton1.Text = 'Find'
$form1.AcceptButton = $findButton1
$form1.Controls.Add($findButton1)

$exportButton1 = New-Object System.Windows.Forms.Button
$exportButton1.Location = New-Object System.Drawing.Point(238, 135)
$exportButton1.Size = New-Object System.Drawing.Size(75, 23)
$exportButton1.Text = 'Export'
$form1.Controls.Add($exportButton1)

$clearButton1 = New-Object System.Windows.Forms.Button
$clearButton1.Location = New-Object System.Drawing.Point(238, 170)
$clearButton1.Size = New-Object System.Drawing.Size(75, 23)
$clearButton1.Text = 'Clear'
$form1.Controls.Add($clearButton1)

$exitButton1 = New-Object System.Windows.Forms.Button
$exitButton1.Location = New-Object System.Drawing.Point(238, 205)
$exitButton1.Size = New-Object System.Drawing.Size(75, 23)
$exitButton1.Text = 'Exit'
$exitButton1.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form1.Controls.Add($exitButton1)

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(20, 20)
$label1.Size = New-Object System.Drawing.Size(280, 20)
$label1.Text = 'Enter the servers below (1 per line):'
$form1.Controls.Add($label1)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(20, 40)
$textBox1.Size = New-Object System.Drawing.Size(200, 250)
$textBox1.Multiline = $true
$textBox1.ScrollBars = 'Vertical'
$textBox1.AcceptsReturn = $true
$form1.Controls.Add($textBox1)
$form1.Add_Shown({ $textBox1.Select() })

$listView1 = New-Object System.Windows.Forms.ListView
$listView1.Location = New-Object System.Drawing.Point(330, 40)
$listView1.Size = New-Object System.Drawing.Size(600, 250)
$listView1.View = 'Details'
$listView1.Columns.Add('Server') | Out-Null
$listView1.Columns.Add('Collection') | Out-Null
$listView1.Columns[1].Width = 500

$outputLabel1 = New-Object System.Windows.Forms.Label
$outputLabel1.Location = New-Object System.Drawing.Point(20, 310)
$outputLabel1.Size = New-Object System.Drawing.Size(800, 40)
$outputLabel1.MaximumSize = New-Object System.Drawing.Size(800, 40)
$outputLabel1.AutoSize = $true
$outputLabel1.Text = 'Welcome to Auto MECM: Find Server Collection'
$form1.Controls.Add($outputLabel1)

# This is the custom comparer class string
# copied from the MSDN article
$comparerClassString = @'
  using System;
  using System.Windows.Forms;
  using System.Drawing;
  using System.Collections;

  public class ListViewItemComparer : IComparer
  {
    private int col;
    public ListViewItemComparer()
    {
      col = 0;
    }
    public ListViewItemComparer(int column)
    {
      col = column;
    }
    public int Compare(object x, object y)
    {
      return String.Compare(
        ((ListViewItem)x).SubItems[col].Text, 
        ((ListViewItem)y).SubItems[col].Text);
    }
  }

'@

# Add the comparer class
Add-Type -TypeDefinition $comparerClassString `
  -ReferencedAssemblies (`
    'System.Windows.Forms', 'System.Drawing')

# Add the event to the ListView ColumnClick event
$columnClick = {
  $listView1.ListViewItemSorter = `
    New-Object ListViewItemComparer($_.Column)
}

$listView1.Add_ColumnClick($columnClick)
$form1.Controls.Add($listView1)

# Add functionality to Find button
$findButton1.add_Click({
    if ($textBox1.TextLength -ne 0) {
      $serverInput = $textBox1.Lines | Where-Object { $_ } | Select-Object -Unique
    }
    else {
      $outputLabel1.ForeColor = [Drawing.Color]::Red
      $outputLabel1.Text = 'ERROR: Server name cannot be left blank'
      Write-Log -Message 'ERROR: Auto-MECM Find Server Collection - textbox server entry was left blank' -Level DEBUG
      return
    }

    try {
      $outputLabel1.ForeColor = [Drawing.Color]::Black
      $outputLabel1.Text = 'Finding the servers collections. Please wait...'

      $listView1.Items.Clear()

      $props = [ordered]@{ 
        Server     = [string]
        Collection = [string]
      }
      
      $Global:arrObj = @()

      foreach ($server in $serverInput) {
        # Get servers collection
        $serverCollection = (Get-WmiObject -ComputerName $SITESERVER -Namespace root/SMS/site_$SITECODE -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$server' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID") | Where-Object { $_.ObjectPath -eq $SCHEDULECOLLECTION } 

        foreach ($item in $serverCollection) {
          # Add results to list view  
          $newItem = New-Object System.Windows.Forms.ListViewItem
          $newItem.Name = $server
          $newItem.Text = $server
          $newItem.Subitems.Add($item.Name)
          $listView1.Items.Add($newItem)

          # Add items for exporting
          $reportObj = New-Object -TypeName PSObject -Property $props
          $reportObj.Server = $server
          $reportObj.Collection = $item.Name
          $Global:arrObj += $reportObj
        }
      }

      # Resize columns if present and according to longest results
      if ($listView1.Items.Count -ne 0) {
        $listView1.AutoResizeColumns(2)
      }

      $serverCount = ($serverInput | Measure-Object).Count
      $outputLabel1.ForeColor = [Drawing.Color]::Black
      $outputLabel1.Text = "Found $($listView1.Items.Count) collections for $serverCount servers"
    }
    catch {
      $outputLabel1.ForeColor = [Drawing.Color]::Red
      $outputLabel1.Text = 'ERROR: Something went wrong'
      Write-Log -Message 'Auto-MECM: Find Server Collection - catch, something went wrong' -Level DEBUG
    }
  })

# Add functionality to Export button
$exportButton1.add_Click({
    if ($listView1.Items.Count -eq 0) {
      $outputLabel1.ForeColor = [Drawing.Color]::Red
      $outputLabel1.Text = 'ERROR: Nothing to export'
      Write-Log -Message 'ERROR: Auto-MECM Find Server Collection - listview is empty, nothing to export' -Level DEBUG
      return
    }
    
    try {
      $outputLabel1.ForeColor = [Drawing.Color]::Black
      $outputLabel1.Text = 'Exporting results to CSV. Please wait...'

      $date = Get-Date -Format yyyy-MM-dd_HH-mm
      $Global:arrObj | Export-Csv -Path ('FileSystem::' + "$logDir\Auto-MECM_FindServerCollection_$date.csv") -NoTypeInformation
      
      $outputLabel1.ForeColor = [Drawing.Color]::Green
      $outputLabel1.Text = "Successfully exported the results to $logDir"
      Write-Log -Message 'ERROR: Auto-MECM Find Server Collection - Successfully exported results' -Level DEBUG
    }
    catch {
      $outputLabel1.ForeColor = [Drawing.Color]::Red
      $outputLabel1.Text = 'ERROR: Failed to export'
      Write-Log -Message 'ERROR: Auto-MECM Find Server Collection - failed to export' -Level DEBUG
    }
  })

# Add functionality to Clear button
$clearButton1.add_Click({
    $textBox1.Clear()
    $listView1.Items.Clear()

    $outputLabel1.ForeColor = [Drawing.Color]::Black
    $outputLabel1.Text = 'Cleared all items'
  })

$formResponse = $form1.ShowDialog()

if ($formResponse -ne [System.Windows.Forms.DialogResult]::OK) {
  break
}