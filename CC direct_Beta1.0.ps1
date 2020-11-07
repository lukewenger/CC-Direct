function printLog {
      $Script:InputBoxProjectName = $InputBoxProjectName.text
      $script:scanPath = $InputBoxScanPath.Text
      $logname = $Script:InputBoxProjectName
      $logName += "_LogCCDirect_"
      $logName += $env:computername 
      $pathLog = "$script:scanpath\$logName.txt"
      $logOutPut = $Script:LogGlobal
      $logOutPut | Out-File $pathLog 
      $logOutPut = $null
}

function processAll  {
        
    foreach ($Scan in $ScanaLL) {
      if ($null -ne $Scan){
            $scan = $ScanaLL[$i]
            $Script:scanAktuell = $Scan.Name
            $Script:answer = 2
            $ccCommand = $ccCommandPart1 
            $cccommand += '"'+$scan+'"'
                              
            $cccommand += $ccCommandPart2
            $CCcommand += $CCSave

            Set-Location C:\Program" "Files\CloudCompare\
         
            Invoke-Expression $CCcommand

            Set-Location $scanPath
            $i = ($i + 1)
            }
      $Script:LogGlobal += Get-Date
      $Script:LogGlobal += ": "
      $Script:LogGlobal += "
      Global shift will be applied: 
      East:     $axisX 
      North:    $axisY
      Height:   $axisZ
      These scans are all being processed:

      $Script:Scanname

      Following operations will be done on all scans:

      $Script:LogOperations
      
      "  
            }
 
 }
function continueFile{
      [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")      
      [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
      $Formcontinue = New-Object System.Windows.Forms.Form
      $Formcontinue.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
      $Formcontinue.Text = "Work With Scans"   
      $Formcontinue.Size = New-Object System.Drawing.Size(600,680)  
      $Formcontinue.StartPosition = "CenterScreen" #loads the window in the center of the screen
      $Formcontinue.BackgroundImageLayout = "Zoom"
      $Formcontinue.MinimizeBox = $False
      $Formcontinue.MaximizeBox = $False
      $Formcontinue.WindowState = "Normal"
      $Formcontinue.SizeGripStyle = "Hide"
      $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
      $Formcontinue.Icon = $Icon
       ###################### END BUTTONS ######################################################
      $recent = $Script:Scanname[$i]
      #### Output Box Field ###############################################################
      $outputBoxContinue = New-Object System.Windows.Forms.RichTextBox
      $outputBoxContinue.Location = New-Object System.Drawing.Size(10,150) 
      $outputBoxContinue.Size = New-Object System.Drawing.Size(500,350)
      $outputBoxContinue.Font = New-Object System.Drawing.Font("Console", 8 ,[System.Drawing.FontStyle]::Regular)
      $outputBoxContinue.MultiLine = $True
      $outputBoxContinue.ScrollBars = "Vertical"
      $outputBoxContinue.Text = "
      Global shift will be applied: 
      East:     $axisX 
      North:    $axisY
      Height:   $axisZ
      recent scan to process to Process:

      $recent

      Following operations will be done:

      $Script:LogOperations
      
      This actions result in the following
      CloudCompare Command:
      
      $ccCommand
   
      "         
      $Script:LogContinue = $outputBoxContinue.Text
      $Formcontinue.Controls.Add($outputBoxContinue)
      $label = New-Object System.Windows.Forms.Label
      $label.Location = New-Object System.Drawing.Point(10,20)
      $label.Size = New-Object System.Drawing.Size(280,20)
      $label.Text = 'Scans to work with'
      $Formcontinue.Controls.Add($label)
      
      $Script:listboxContinue = New-Object System.Windows.Forms.ListBox
      $Script:listboxContinue.Location = New-Object System.Drawing.Point(10,60)
      $Script:listboxContinue.Size = New-Object System.Drawing.Size(350,70)
      $Script:listboxContinue.SelectionMode = 'MultiExtended'  
      
      $Script:LogGlobal += "With a little help by Boxxer, CloudCompare has performed following Operations and task for you:
       "
      $li = 0
      foreach ($element in $Script:Scanname)
      {
      $element = $Script:Scanname[$li]
      [void] $Script:listboxContinue.Items.Add($element)
      $Formcontinue.Controls.Add($Script:listboxContinue)
      $li = ($li + 1)
      }
      
      
      $Formcontinue.Topmost = $false

      $RunCommand = New-Object System.Windows.Forms.Button
      $RunCommand.Location = New-Object System.Drawing.Size(10,540)
      $RunCommand.Size = New-Object System.Drawing.Size(150,60)
      $RunCommand.Text = "Run Command"
      $RunCommand.Add_Click({
            $Script:answer = 0
                  ###Log if answer = 0 (Process scan)
      if ($script:answer -eq 0) {
            $Script:LogGlobal += Get-Date
            $Script:LogGlobal += ": "
            $Script:LogGlobal += $Script:LogContinue          
            }
      $Formcontinue.close()})
      $RunCommand.Cursor = [System.Windows.Forms.Cursors]::Hand
      $Formcontinue.Controls.Add($RunCommand)


      $Skip = New-Object System.Windows.Forms.Button
      $Skip.Location = New-Object System.Drawing.Size(160,540)
      $Skip.Size = New-Object System.Drawing.Size(150,60)
      $Skip.Text = "skip"
      $Skip.Add_Click({
            $Script:answer = 1
            $Formcontinue.close()})
      $Skip.Cursor = [System.Windows.Forms.Cursors]::Hand
      $Formcontinue.Controls.Add($Skip)
     
      $loopAll = New-Object System.Windows.Forms.Button
      $loopAll.Location = New-Object System.Drawing.Size(360,65)
      $loopAll.Size = New-Object System.Drawing.Size(150,60)
      $loopAll.Text = "loopAll"
      $loopAll.Add_Click({
            $Script:answer = 4
            processAll
            $Formcontinue.close()})
      $loopAll.Cursor = [System.Windows.Forms.Cursors]::Hand
      $Formcontinue.Controls.Add($loopAll)
     

      $abortAction = New-Object System.Windows.Forms.Button
      $abortAction.Location = New-Object System.Drawing.Size(310,540)
      $abortAction.Size = New-Object System.Drawing.Size(150,60)
      $abortAction.Text = "Abort"
      $abortAction.Add_Click({
            $Script:answer = 2
            $Formcontinue.close()})
      $abortAction.Cursor = [System.Windows.Forms.Cursors]::Hand
      $Formcontinue.Controls.Add($abortAction)
      

      $Script:processingScans = $Script:listboxContinue.SelectedItems
      ##############################################
   
      $Formcontinue.Add_Shown({$Formcontinue.Activate()})
      [void] $Formcontinue.ShowDialog()
}

function RunCommand {
            ###assign GUI input    
    $script:scanPath = $InputBoxScanPath.Text 
    $axisX = $InputBoxAxisX.Text
    $axisY = $InputBoxAxisY.Text
    $axisZ = $InputBoxAxisZ.Text
    $SubSampleTo = $inputBoxSS.text
    $neighbourCount = $InputBoxNN.text
    $overlap = $InputBoxOL.text
    $octree = $inputBoxOC.text
    $nsigma = $InputBoxNsigma.text
    
    $SubSampleTo = [decimal]$SubSampleTo
    $neighbourCount = [decimal]$neighbourCount
    $overlap = [decimal]$overlap
    $octree = [decimal]$octree
    $nsigma = [decimal]$nsigma

    $axisX = [decimal]$axisX
    $axisY = [decimal]$axisY
    $axisZ = [decimal]$axisZ
    $Script:axisZ = $axisZ*-1
    $Script:axisY = $axisY*-1
    $Script:axisX = $axisX*-1


    #Define CloudCompareSnippets
    $Script:LogOperations = $null

      $CCnormComp = " -COMPUTE_NORMALS"
      $CCnormOrient = " -ORIENT_NORMS-MST $neighbourCount"
      $CCconnectedcomponent = " -EXTRACT_CC $octree $neighbourCount"
      $CCc2c = " -C2C_DIST"
      $CCmerge = " -MERGE_CLOUDS"
      $CCSubsample = " -ss SPATIAL $SubSampleTo"
      $CCsor =" _SOR $neighbourCount $nsigma"
      $CCicp = " -ICP -OVERLAP $overlap -FARTHEST_REMOVAL"
      $noTimeStamp = " -NO_TIMESTAMP"
      $CCdropGS = " -DROP_GLOBAL_SHIFT"
      $removeScalar = " -REMOVE_ALL_SFS"
      $CCSave = " -SAVE_CLOUDS"
      $CCSaveAllAtOnce = " All_AT_ONCE"

      #assign Path
      $ScanAll = (Get-ChildItem -Recurse $scanPath\ -Include *.ply, *.laz, *.e57, *.pts, *.las, *.xyz, *.csv, *.bin)
 
      $Script:scanName = $ScanAll.Name    
     
      $i = 0
                  
     
      #CloudCompare Options
      if ($listboxCCops.selectedItems -eq 'Combine Scans in CC Project'){
            $CCSave += $CCSaveAllAtOnce
            $Script:LogOperations += "Scans will be saved in one CloudCompare Projekt"
            $Script:DynamicListBox
            }
      

      if ($listboxCCops.selectedItems -eq 'Subsample Scans'){
            $ccCommandPart2 += $CCSubsample
            $Script:LogOperations += "Scans will be subsampled
            "
            }       

      if ($listboxCCops.selectedItems -eq 'ICP Alignement'){
            $ccCommandPart2 += $CCicp
            $Script:LogOperations += "Scans will be aligned
            "
            }       
      
      if ($listboxCCops.selectedItems -eq 'C2C Distance'){
            $ccCommandPart2 += $CCc2c
            $Script:LogOperations += "Distance between scans will be compared
            "
            }    

      if ($listboxCCops.selectedItems -eq 'SOR Filter'){
            $ccCommandPart2 += $CCsor
            $Script:LogOperations += "Statistical outlier removal filter is applied
            "
            }
      
      if ($listboxCCops.selectedItems -eq 'Labeled Connected Component'){
            $ccCommandPart2 += $CCconnectedcomponent
            $Script:LogOperations += "Labeled connected components are being extracted
            "
            }
      
      if ($listboxCCops.selectedItems -eq 'Compute Normals'){
            $ccCommandPart2 += $CCnormComp
            $Script:LogOperations += "Normals are computed
            "
            }   
      
      if ($listboxCCops.selectedItems -eq 'Orient Normals'){
            $ccCommandPart2 += $CCnormOrient
            $Script:LogOperations += "Normals are oriented
            "
            }
            

            
      $ccCommandPart1 = "./cloudcompare -auto_save OFF -O -GLOBAL_SHIFT ($Script:axisX,$Script:axisY,$Script:axisZ) "
      $Script:openFile = " -O -GLOBAL_SHIFT ($Script:axisX,$Script:axisY,$Script:axisZ) "
            
      ###define Export Format Actions

      if ($listboxExForm.selectedItem -eq "Bin"){
            $exFormat = " "
            $ccCommandPart2 += $exFormat
            $Script:LogOperations += "Export as CloudCompare Native Bin File
            "
            }
      if ($listboxExForm.SelectedItem -eq "e57"){
            $exFormat = " -C_EXPORT_FMT E57"
            $ccCommandPart2 += $exFormat
            $Script:LogOperations += "Export as e57
            "
            }
      
      if ($listboxExForm.SelectedItem -eq "Ply binary"){
            $exFormat = " -PLY_EXPORT_FMT ASCII"
            $ccCommandPart2 += $exFormat
            $Script:LogOperations += "Export as Ply binary
            "
            }
      if ($listboxExForm.SelectedItem -eq "las"){
            $exFormat = " -C_EXPORT_FMT LAS"
            $ccCommandPart2 += $exFormat
            $Script:LogOperations += "Export as las
            "
            }
                    
      # Export Options
      
      if ($listboxCCExport.selectedItems -eq "No Timestamp on save"){
        $ccCommandPart2 += $noTimeStamp
        $Script:LogOperations += "No Timestamp added
        "
        }
      
      if ($listboxCCExport.selectedItems -eq 'Merge Scans'){
            $ccCommandPart2 += $CCmerge
            $Script:LogOperations += "On Export all scans will be merged
            "
      }

      if ($listboxCCExport.selectedItems -eq 'drop global shift on save'){
            $ccCommandPart2 += $CCdropGS
            $Script:LogOperations += "DANGER DANGER
            ALL GLOBAL SHIFT WILL BE DROPPED THIS INSTANT, 
            DO YOU REALLY NOW WHAT YOU ARE DOING??? 
            DANGER DANGER
            "
      }       

      if ($listboxCCExport.selectedItems -eq 'remove all scalar fields'){
            $ccCommandPart2 += $removeScalar
            $Script:LogOperations += "All scalar fields will be used
            "
      }
      
      ##loop through files

      foreach ($Scan in $ScanaLL) {
            if ($null -ne $Scan){
                  $scan = $ScanaLL[$i]
                  $Script:scanAktuell = $Scan.Name
                  $Script:answer = 2
                  $ccCommand = $ccCommandPart1 
                  $cccommand += '"'+$scan+'"'
                  $Script:log += $Script:LogGlobal
            
                  if ($listboxCCops.selectedItems -eq 'Combine Scans in CC Project'){
                    $i = 1    
                    foreach ($scan in $ScanAll) {
                        if ($null -eq $ScanAll[$i]){continue}
                        $scan = $ScanAll[$i]
                        $scanCombine += $openFile+' "'+$scan+'"'
                        $i = $i+1
                        }
                        $ccCommand += $scanCombine
                        $cccommand += $ccCommandPart2
                        $CCcommand += $CCSave                                                
                        $Script:LogOperations += "Scans will be saved in one CloudCompare Projekt"
                        $Script:log += " 
                        The following scans are being processed            
                        $scanAll
                        Global Shift settings for processing:
                        X East:     $axisX 
                        Y North:    $axisY
                        Z Height:   $axisZ
      
                        Operations apllied:
                        $Script:LogOperations
                        "
                        $outputBox.text = $outputBox.Text.insert(0, $Script:LogGlobal)
                        continueFile
                        Set-Location C:\Program" "Files\CloudCompare\    
                        Invoke-Expression $CCcommand
            
                        Set-Location $scanPath
                        exit       
                        }
                        
                  $cccommand += $ccCommandPart2
                  $CCcommand += $CCSave

                  Set-Location C:\Program" "Files\CloudCompare\

                  continueFile
                  if ($Script:answer -eq 0){$outputBox.text = $outputBox.Text.insert(0, $Script:LogGlobal)}
                  if ($Script:answer -eq 1){$i = $i + 1 
                        continue}
                  if ($Script:answer -eq 2){exit}


                 
                  Invoke-Expression $CCcommand
      
                  Set-Location $scanPath
                  $i = ($i + 1)
                  }

      }
      
}

function MainGui {
    ###################### CREATING PS GUI TOOL #############################
   
      #### Form settings #################################################################
      [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
      [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
   
      $Form = New-Object System.Windows.Forms.Form
      $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
      $Form.Text = "Work With Scans"   
      $Form.Size = New-Object System.Drawing.Size(1010,680)  
      $Form.StartPosition = "CenterScreen" #loads the window in the center of the screen
      $Form.BackgroundImageLayout = "Zoom"
      $Form.MinimizeBox = $False
      $Form.MaximizeBox = $False
      $Form.WindowState = "Normal"
      $Form.SizeGripStyle = "Hide"
      $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
      $Form.Icon = $Icon
   
      #### Title - Powershell GUI Tool ###################################################
      $Label = New-Object System.Windows.Forms.Label
      $LabelFont = New-Object System.Drawing.Font("Calibri",18,[System.Drawing.FontStyle]::Bold)
      $Label.Font = $LabelFont
      $Label.Text = "Boxxers little helper for CloudCompare"
      $Label.AutoSize = $True
      $Label.Location = New-Object System.Drawing.Size(300,40) 
      $Form.Controls.Add($Label)

   
      #### Input window with "Scanpath ##########################################

      $InputBoxScanPath = New-Object System.Windows.Forms.TextBox 
      $InputBoxScanPath.Location = New-Object System.Drawing.Size(10,100) 
      $InputBoxScanPath.Size = New-Object System.Drawing.Size(180,20) 
      $Form.Controls.Add($InputBoxScanPath)
      
      $LabelScanPath = New-Object System.Windows.Forms.Label
      $LabelScanPath.Text = "Scanpath:"
      $LabelScanPath.AutoSize = $True
      $LabelScanPath.Location = New-Object System.Drawing.Size(15,80) 
      $form.Controls.Add($LabelScanPath)

 #### Input window with "Project Name" ##########################################

 $InputBoxProjectName = New-Object System.Windows.Forms.TextBox 
 $InputBoxProjectName.Location = New-Object System.Drawing.Size(10,50) 
 $InputBoxProjectName.Size = New-Object System.Drawing.Size(180,20) 
 $Form.Controls.Add($InputBoxProjectName)
 
 $LabelProjectName = New-Object System.Windows.Forms.Label
 $LabelProjectName.Text = "ProjectName:"
 $LabelProjectName.AutoSize = $True
 $LabelProjectName.Location = New-Object System.Drawing.Size(15,30) 
 $form.Controls.Add($LabelProjectName)

    ###########Input Global Shift
##title      
      $LabelGS = New-Object System.Windows.Forms.Label
      $LabelGS.Text = "Global Shift:"
      $LabelGS.AutoSize = $True
      $LabelGS.Location = New-Object System.Drawing.Size(15,140) 
      $form.Controls.Add($LabelGS)
##Axis X
$LabelX = New-Object System.Windows.Forms.Label
$LabelX.Text = "X-Achse (East):"
$LabelX.AutoSize = $True
$LabelX.Location = New-Object System.Drawing.Size(15,160) 
$form.Controls.Add($LabelX)

$InputBoxAxisX = New-Object System.Windows.Forms.TextBox 
$InputBoxAxisX.Location = New-Object System.Drawing.Size(15,180) 
$InputBoxAxisX.Size = New-Object System.Drawing.Size(120,20) 
$Form.Controls.Add($InputBoxAxisX)

##Axis Y
$LabelY = New-Object System.Windows.Forms.Label
$LabelY.Text = "Y-Achse (Nord):"
$LabelY.AutoSize = $True
$LabelY.Location = New-Object System.Drawing.Size(15,200) 
$form.Controls.Add($LabelY)

$InputBoxAxisY = New-Object System.Windows.Forms.TextBox 
$InputBoxAxisY.Location = New-Object System.Drawing.Size(15,220) 
$InputBoxAxisY.Size = New-Object System.Drawing.Size(120,20) 
$Form.Controls.Add($InputBoxAxisY)

##Axis Z
$LabelZ = New-Object System.Windows.Forms.Label
$LabelZ.Text = "Z-Achse (Hoehe):"
$LabelZ.AutoSize = $True
$LabelZ.Location = New-Object System.Drawing.Size(15,240) 
$form.Controls.Add($LabelZ)

$InputBoxAxisZ = New-Object System.Windows.Forms.TextBox 
$InputBoxAxisZ.Location = New-Object System.Drawing.Size(15,260) 
$InputBoxAxisZ.Size = New-Object System.Drawing.Size(120,20) 
$Form.Controls.Add($InputBoxAxisZ)
 
##Subsample to
$LabelSS = New-Object System.Windows.Forms.Label
$LabelSS.Text = "subsample to in m:"
$LabelSS.AutoSize = $True
$LabelSS.Location = New-Object System.Drawing.Size(15,300) 
$form.Controls.Add($LabelSS)

$InputBoxSS = New-Object System.Windows.Forms.TextBox 
$InputBoxSS.Location = New-Object System.Drawing.Size(10,320) 
$InputBoxSS.Size = New-Object System.Drawing.Size(180,20) 
$Form.Controls.Add($InputBoxSS)


## Number of Neighbours

$LabelNN = New-Object System.Windows.Forms.Label
$LabelNN.Text = "min. number of neighbours:"
$LabelNN.AutoSize = $True
$LabelNN.Location = New-Object System.Drawing.Size(15,360) 
$form.Controls.Add($LabelNN)

$InputBoxNN = New-Object System.Windows.Forms.TextBox 
$InputBoxNN.Location = New-Object System.Drawing.Size(10,380) 
$InputBoxNN.Size = New-Object System.Drawing.Size(180,20) 
$Form.Controls.Add($InputBoxNN)

##### Octree Level

$LabelOctree = New-Object System.Windows.Forms.Label
$LabelOctree.Text = "Level of Octree:"
$LabelOctree.AutoSize = $True
$LabelOctree.Location = New-Object System.Drawing.Size(15,400) 
$form.Controls.Add($LabelOctree)

$InputBoxOC = New-Object System.Windows.Forms.TextBox 
$InputBoxOC.Location = New-Object System.Drawing.Size(10,420) 
$InputBoxOC.Size = New-Object System.Drawing.Size(180,20) 
$Form.Controls.Add($InputBoxOC)

### approx. overlap

$LabelOverlap = New-Object System.Windows.Forms.Label
$LabelOverlap.Text = "approximate overlap:"
$LabelOverlap.AutoSize = $True
$LabelOverlap.Location = New-Object System.Drawing.Size(15,460) 
$form.Controls.Add($LabelOverlap)

$InputBoxOL = New-Object System.Windows.Forms.TextBox 
$InputBoxOL.Location = New-Object System.Drawing.Size(10,480) 
$InputBoxOL.Size = New-Object System.Drawing.Size(180,20) 
$Form.Controls.Add($InputBoxOL)

##Standard deviation Nsigma

$LabelNsigma = New-Object System.Windows.Forms.Label
$LabelNsigma.Text = "Nsigma:"
$LabelNsigma.AutoSize = $True
$LabelNsigma.Location = New-Object System.Drawing.Size(15,520) 
$form.Controls.Add($LabelNsigma)

$InputBoxNsigma = New-Object System.Windows.Forms.TextBox 
$InputBoxNsigma.Location = New-Object System.Drawing.Size(10,540) 
$InputBoxNsigma.Size = New-Object System.Drawing.Size(180,20) 
$Form.Controls.Add($InputBoxNsigma)
 

     ##Add Listbox CC Commands

    $form.Controls.Add($listboxCCops)
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(200,120)
    $label.Size = New-Object System.Drawing.Size(280,18)
    $label.Text = 'CC Commands Basic:'
    $form.Controls.Add($label)

   $listboxCCops = New-Object System.Windows.Forms.listbox
   $listboxCCops.Location = New-Object System.Drawing.Point(200,140)
   $listboxCCops.Size = New-Object System.Drawing.Size(260,250)
   
   $listboxCCops.SelectionMode = 'MultiExtended'
   
   [void] $listboxCCops.Items.Add('Subsample Scans')
   [void] $listboxCCops.Items.Add('Compute Normals')
   [void] $listboxCCops.Items.Add('Orient Normals')
   [void] $listboxCCops.Items.Add('SOR Filter')
   [void] $listboxCCops.Items.Add('Labeled Connected Component')
   [void] $listboxCCops.Items.Add('C2C Distance')
   [void] $listboxCCops.Items.Add('ICP Alignement')
   [void] $listboxCCops.Items.Add('Combine Scans in CC Project')
   

   $form.Controls.Add($listboxCCops)

   
         ##### Checkbox export
         $label = New-Object System.Windows.Forms.Label
         $label.Location = New-Object System.Drawing.Point(480,100)
         $label.Size = New-Object System.Drawing.Size(280,18)
         $label.Text = 'Export Format:'
         $form.Controls.Add($label)
         
         $listboxExForm = New-Object System.Windows.Forms.ListBox
         $listboxExForm.Location = New-Object System.Drawing.Point(480,120)
         $listboxExForm.Size = New-Object System.Drawing.Size(100,80)
         
   
         
         [void] $listboxExForm.Items.Add('Bin')
         [void] $listboxExForm.Items.Add('e57')
         [void] $listboxExForm.Items.Add('Ply binary')
         [void] $listboxExForm.Items.Add('Las')

     
         $form.Controls.Add($listboxExForm)
         
         $form.Topmost = $false
####Export Options

     ##Add Listbox Export Options

     $form.Controls.Add($listboxCCExport)
     $label = New-Object System.Windows.Forms.Label
     $label.Location = New-Object System.Drawing.Point(600,120)
     $label.Size = New-Object System.Drawing.Size(200,18)
     $label.Text = 'Export Options:'
     $form.Controls.Add($label)
 
    $listboxCCExport = New-Object System.Windows.Forms.ListBox
    $listboxCCExport.Location = New-Object System.Drawing.Point(600,140)
    $listboxCCExport.Size = New-Object System.Drawing.Size(260,250)
    
    $listboxCCExport.SelectionMode = 'MultiExtended'
    
    [void] $listboxCCExport.Items.Add('Merge Scans')
    [void] $listboxCCExport.Items.Add('drop global shift on save')
    [void] $listboxCCExport.Items.Add('remove all scalar fields')
    [void] $listboxCCExport.Items.Add('No Timestamp on save')

     
    $form.Controls.Add($listboxCCExport)
 
  
      #### PrintLog ###################################################################
      $LogPrint = New-Object System.Windows.Forms.Button
      $LogPrint.Location = New-Object System.Drawing.Size(820,400)
      $LogPrint.Size = New-Object System.Drawing.Size(80,30)
      $LogPrint.Text = "Print Log"
      $LogPrint.Add_Click({PrintLog})
      $LogPrint.Cursor = [System.Windows.Forms.Cursors]::Hand
      $form.Controls.Add($LogPrint)
      

    #### RunCommand ###################################################################
      $RunCommand = New-Object System.Windows.Forms.Button
      $RunCommand.Location = New-Object System.Drawing.Size(820,530)
      $RunCommand.Size = New-Object System.Drawing.Size(150,60)
      $RunCommand.Text = "Run Command"
      $RunCommand.Add_Click({RunCommand})
      $RunCommand.Cursor = [System.Windows.Forms.Cursors]::Hand
      $form.Controls.Add($RunCommand)


      ###################### END BUTTONS ######################################################
   
      #### Output Box Field ###############################################################
      $LogFinal = "
      Welcome
Choose your commands
"
      $LogFinal += $Script:LogGlobal
      $outputBox = New-Object System.Windows.Forms.RichTextBox
      $outputBox.Location = New-Object System.Drawing.Size(300,400) 
      $outputBox.Size = New-Object System.Drawing.Size(500,200)
      $outputBox.Font = New-Object System.Drawing.Font("Console", 8 ,[System.Drawing.FontStyle]::Regular)
      $outputBox.MultiLine = $True
      $outputBox.ScrollBars = "Vertical"
      $outputBox.Text = $LogFinal
      $Form.Controls.Add($outputBox)
   
      ##############################################
   
      $Form.Add_Shown({$Form.Activate()})
      [void] $Form.ShowDialog()
    }
      MainGui