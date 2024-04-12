# Laden der Assembly für Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Funktion zur Erstellung des Hauptfensters
Function MainWindow {
    Begin {
        #region Variablen Initialisierung
        # Hauptfenster Variable
        [Object] $mainWindow = $null

        # Initialisierung der Panel (Abschnitts-) Variablen
        [Object] $panelMailFrom                  = $null
        [Object] $panelMailTo                    = $null
        [Object] $panelMailCC                    = $null
        [Object] $panelSubject                   = $null
        [Object] $panelMailText                  = $null
        [Object] $panelMailTextSalutation        = $null
        [Object] $panelMailTextDevices           = $null
        [Object] $panelMailTextEnvironment       = $null
        [Object] $panelMailTextDate              = $null
        [Object] $panelMailTextSecurityPatch     = $null
        [Object] $panelMailTextReboot            = $null
        [Object] $panelMailTextOtherInformations = $null
        [Object] $panelMailTextBWZInform         = $null
        [Object] $panelMailTextResponsible       = $null
        [Object] $panelMailTextEnd               = $null
        
        # E-Mail-Variablen Von
        [Object] $mailFromLabel      = $null
        [Object] $mailFromMailAdress = $null

        # E-Mail Variablen An
        [Object] $mailToLabel            = $null
        [Object] $mailToCheckedListBox   = $null

        # E-Mail Variablen weitere Empfänger
        [Object] $mailCCLabel   = $null
        [Object] $mailCCTextBox = $null

        # E-Mail Variablen Betreff
        [Object] $mailSubjectLabel       = $null
        [Object] $mailSubjectPackageName = $null

        # E-Mail Text Variablen
        [Object] $mailTextLabel                        = $null
        [Object] $mailTextSalutationLabel              = $null
        [Object] $mailTextIntroductionLabel            = $null
        [Object] $mailTextDeviceLabel                  = $null
        [Object] $mailTextAllDevicesComboBox           = $null
        [Object] $mailTextDevicesComboBox              = $null
        [Object] $mailTextEnvironmentLabel             = $null
        [Object] $mailTextEnvironmentListBox           = $null
        [Object] $mailTextDateLabel                    = $null
        [Object] $mailTextDateSelector                 = $null
        [Object] $mailTextDateDeadLineLabel            = $null
        [Object] $mailTextDateDeadLineYesComboBox      = $null
        [Object] $mailTextDateDeadLineNoComboBox       = $null
        [Object] $mailTextSecurityPatchLabel           = $null
        [Object] $mailTextSecurityPatchYesLabel        = $null
        [Object] $mailTextPatchRoutineLabel            = $null
        [Object] $mailTextPatchRoutineTypeListBox      = $null
        [Object] $mailTextRebootLabel                  = $null
        [Object] $mailTextRebootYesCheckBox            = $null
        [Object] $mailTextRebootNoCheckBox             = $null
        [Object] $mailTextOtherInformationsLabel       = $null
        [Object] $mailTextOtherInformationsRichTextBox = $null
        [Object] $mailTextUserInfoLabel                = $null
        [Object] $mailTextBWZLabel                     = $null
        [Object] $mailTextEndLabel                     = $null
        
        # Schriftarten Variablen
        [Object] $standradfont       = $null
        [Object] $standradfontBold   = $null

        # Hauptfenster Variablen für Größe und Position
        [Int]    $mainWindowWidth    = 1024
        [Int]    $mainWindowHeight   = 1024
        [Int]    $mainWindowX        = 0
        [Int]    $mainWindowY        = 0

        # Listen Variablen
        [Array]  $mailToList          = @()
        [Array]  $mailEnvironmentList = @()
        [Array]  $mailInstallTypeList = @()
        [Array]  $mailResponsible     = @()
        [XML]    $mailAdresses        = ""
        #endregion

        # Berechnung an welcher Bildschirm Position das Hauptfenster angezeigt werden soll
        $mainWindowX = $( ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width/2) - ($mainWindowWidth/2) )
        $mainWindowY = $( ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height/2) - ($mainWindowHeight /2) )

        # Festlegen der Schrifarten
        $standradfont     = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Regular)
        $standradfontBold = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)

        #Festlegen der Listen
        $mailToList          = ("Alle Kunden","BADV","BAMF","BBK","BeschA","BKGE","BMWSB","BVA","BZSt","GWK","ITZBund","KVDB","TBC","ZOLL mit BWZ","ZOLL ohne BWZ")
        $mailEnvironmentList = ("ALLE", "VZD", "ZVD-P1", "ZVD-P2")
        $mailInstallTypeList = ("Ablösekette", "Updatemechanismus durch die Herstellerroutine", "Parallele Installation")
        $mailResponsible     = ("Marc Leonhard", "Partik Leonhard", "Sven Wendt", "Thomas Fischer")
    }

    Process {
        #region Hauptfenster
        # Erstellen des Hauptfensters
        $mainWindow               = New-Object System.Windows.Forms.Form
        $mainWindow.Location      = New-Object System.Drawing.Point($mainWindowX, $mainWindowY)
        $mainWindow.Size          = New-Object System.Drawing.Size($mainWindowWidth, $mainWindowHeight)
        $mainWindow.Text          = 'Ankündigungsmail'
        $mainWindow.StartPosition = 'Manual'
        $mainWindow.MinimumSize   = New-Object System.Drawing.Size($mainWindowWidth, $mainWindowHeight)

        # Wenn ALT+F4 gedrückt wird, soll das Hauptfenster geschlossen werden
        $mainWindow.Add_KeyDown({
            If(($_.Alt -eq $true) -and ($_.KeyCode -eq "F4"))
            {
                $mainWindow.Close()
                $mainWindow.Dispose()
            }
        })
        #endregion

        #region MailFrom
        # Erstellen eines Panels
        $panelMailFrom          = New-Object System.Windows.Forms.Panel
        $panelMailFrom.Location = New-Object System.Drawing.Point(1,1)
        $panelMailFrom.Size     = New-Object System.Drawing.Size(1006,30)
        $panelMailFrom.BorderStyle = "Fixed3D"

         # Erstellen einer MailVon Label
        $mailFromLabel          = New-Object System.Windows.Forms.Label
        $mailFromLabel.Location = New-Object System.Drawing.Point(10,5)
        $mailFromLabel.Size     = New-Object System.Drawing.Size(40,20)
        $mailFromLabel.Text     = "Von:"
        $mailFromLabel.Font     = $standradfontBold

        # Erstellen einer Mail-Von
        $mailFromMailAdress          = New-Object System.Windows.Forms.Label
        $mailFromMailAdress.Location = New-Object System.Drawing.Point(120,5)
        $mailFromMailAdress.Size     = New-Object System.Drawing.Size(200,20)
        $mailFromMailAdress.Text     = "FB-Rollout (ITZBund)"
        $mailFromMailAdress.Font     = $standradfont
        
        $panelMailFrom.Controls.Add($mailFromLabel)
        $panelMailFrom.Controls.Add($mailFromMailAdress)

        #endregion

        #region MailTo
        $panelMailTo          = New-Object System.Windows.Forms.Panel
        $panelMailTo.Location = New-Object System.Drawing.Point(1,31)
        $panelMailTo.Size     = New-Object System.Drawing.Size(1006,85)
        $panelMailTo.BorderStyle = "Fixed3D"

        # Erstellen einer MailTo-ComboBox
        $mailToLabel          = New-Object System.Windows.Forms.Label
        $mailToLabel.Location = New-Object System.Drawing.Point(10,30)
        $mailToLabel.Size     = New-Object System.Drawing.Size(40,20)
        $mailToLabel.Text     = "An:"
        $mailToLabel.Font     = $standradfontBold

        # Erstellen einer CheckedListBox
        $mailToCheckedListBox              = New-Object System.Windows.Forms.CheckedListBox
        $mailToCheckedListBox.Location     = New-Object System.Drawing.Point(120,5)
        $mailToCheckedListBox.Size         = New-Object System.Drawing.Size(200,80)
        $mailToCheckedListBox.Font         = $standradfont
        $mailToCheckedListBox.CheckOnClick = $true

        $mailToCheckedListBox.Add_SelectedValueChanged({
            $checked = 0
            
            For($i=1; $i -lt $mailToCheckedListBox.Items.Count; $i++)
            {
                If( $mailToCheckedListBox.GetItemCheckState($i) -eq "Checked")
                {
                    $checked++
                }   
            }

            If($checked -eq ($mailToCheckedListBox.Items.Count - 1) )
            {
                 for ($i = 1; $i -lt $mailToCheckedListBox.Items.Count; $i++)
                 {
                    $mailToCheckedListBox.SetItemChecked($i, $false)
                 }
                 $mailToCheckedListBox.SetItemChecked(0, $true)
            }
        })

        $mailToCheckedListBox.Add_Click({
            $selectedItem = $mailToCheckedListBox.SelectedItem

            # Überprüfen, ob Element 1 ausgewählt ist
            if ($selectedItem -eq "Alle Kunden") {
                # Wenn Element 1 ausgewählt ist und ein anderes Element geklickt wird, deaktivieren Sie Element 1
                if ($mailToCheckedListBox.CheckedIndices.Contains(0) -and $mailToCheckedListBox.SelectedIndex -ne 0)
                {
                    $mailToCheckedListBox.SetItemChecked(0, $false)
                }
                # Wenn Element 1 deaktiviert ist und es ausgewählt wird, deaktivieren Sie alle anderen Elemente
                elseif (-not $mailToCheckedListBox.CheckedIndices.Contains(0) -and $mailToCheckedListBox.SelectedIndex -eq 0)
                {
                    for ($i = 1; $i -lt $mailToCheckedListBox.Items.Count; $i++)
                    {
                        $mailToCheckedListBox.SetItemChecked($i, $false)
                    }
                }
            }
            # Wenn ein anderes Element als Element 1 ausgewählt ist, deaktivieren Sie Element 1
            else {
                $mailToCheckedListBox.SetItemChecked(0, $false)
            }
        })

        # Hinzufügen von Items in die CheckedListBox
        for ($i = 0; $i -lt $mailToList.Count; $i++) {
            $mailToCheckedListBox.Items.Add($mailToList[$i]) | Out-Null
        }

        $mailToCheckedListBox.SetItemChecked(0, $true)

        $panelMailTo.Controls.Add($mailToLabel)
        $panelMailTo.Controls.Add($mailToCheckedListBox)

        #endregion

        #region CC
        $panelMailCC             = New-Object System.Windows.Forms.Panel
        $panelMailCC.Location    = New-Object System.Drawing.Point(1,116)
        $panelMailCC.Size        = New-Object System.Drawing.Size(1006,35)
        $panelMailCC.BorderStyle = "Fixed3D"

        $mailCCLabel          = New-Object System.Windows.Forms.Label
        $mailCCLabel.Location = New-Object System.Drawing.Point(10,7)
        $mailCCLabel.AutoSize = $true
        $mailCCLabel.Text     = "CC:"
        $mailCCLabel.Font     = $standradfontBold

        $mailCCTextBox          = New-Object System.Windows.Forms.TextBox
        $mailCCTextBox.Location = New-Object System.Drawing.Point(120,5)
        $mailCCTextBox.Size     = New-Object System.Drawing.Size(880,20)
        $mailCCTextBox.Text     = "Weitere Empfänger"
        $mailCCTextBox.Font     = $standradfont

        $mailCCTextBox.Add_Click({
            $mailCCTextBox.Text = ""
        })

        $mailCCTextBox.Add_LostFocus({
            If( [String]::IsNullOrEmpty($mailCCTextBox.Text) )
            {
                $mailCCTextBox.Text    = "Weitere Empfänger"
            }
        })

        $panelMailCC.Controls.Add($mailCCLabel)
        $panelMailCC.Controls.Add($mailCCTextBox)
        #endregion

        #region Subject
        $panelSubject             = New-Object System.Windows.Forms.Panel
        $panelSubject.Location    = New-Object System.Drawing.Point(1,151)
        $panelSubject.Size        = New-Object System.Drawing.Size(1006,35)
        $panelSubject.BorderStyle = "Fixed3D"

        $mailSubjectLabel          = New-Object System.Windows.Forms.Label
        $mailSubjectLabel.Location = New-Object System.Drawing.Point(10,10)
        $mailSubjectLabel.AutoSize = $true
        $mailSubjectLabel.Text     = "Betreff: Rollout"
        $mailSubjectLabel.Font     = $standradfontBold

        $mailSubjectPackageName          = New-Object System.Windows.Forms.TextBox
        $mailSubjectPackageName.Location = New-Object System.Drawing.Point(120,7)
        $mailSubjectPackageName.Size     = New-Object System.Drawing.Size(880,20)
        $mailSubjectPackageName.Text     = "Paketname"
        $mailSubjectPackageName.Font     = $standradfont

        $mailSubjectPackageName.Add_Click({
            $mailSubjectPackageName.Text = ""
        })

        $mailSubjectPackageName.Add_LostFocus({
            If( [String]::IsNullOrEmpty($mailSubjectPackageName.Text) )
            {
                $mailSubjectPackageName.Text    = "Paketname"
                $mailTextIntroductionLabel.Text = 'das Team Rollout des ITZBunds wurde beauftragt, das Softwarepaket "'+$($mailSubjectPackageName.Text)+'" zu verteilen.'
            }
            Else
            {
                $mailTextIntroductionLabel.Text = 'das Team Rollout des ITZBunds wurde beauftragt, das Softwarepaket "'+$($mailSubjectPackageName.Text)+'" zu verteilen.'
            }
        })

        $panelSubject.Controls.Add($mailSubjectLabel)
        $panelSubject.Controls.Add($mailSubjectPackageName)
        
        #endregion

        #region Mail-Text
        $panelMailText             = New-Object System.Windows.Forms.Panel
        $panelMailText.Location    = New-Object System.Drawing.Point(1,186)
        $panelMailText.Size        = New-Object System.Drawing.Size(1006,800)
        $panelMailText.BorderStyle = "Fixed3D"

        $mailTextLabel          = New-Object System.Windows.Forms.Label
        $mailTextLabel.Location = New-Object System.Drawing.Point(10,20)
        $mailTextLabel.AutoSize = $true
        $mailTextLabel.Text     = "E-Mail-Text:"
        $mailTextLabel.Font     = $standradfontBold

        $panelMailText.Controls.Add($mailTextLabel)
        #endregion

        #region Mail-Text-Anrede
        $panelMailTextSalutation             = New-Object System.Windows.Forms.Panel
        $panelMailTextSalutation.Location    = New-Object System.Drawing.Point(0,40)
        $panelMailTextSalutation.Size        = New-Object System.Drawing.Size(1002,55)
        $panelMailTextSalutation.BorderStyle = "Fixed3D"

        $mailTextSalutationLabel          = New-Object System.Windows.Forms.Label
        $mailTextSalutationLabel.Location = New-Object System.Drawing.Point(10,5)
        $mailTextSalutationLabel.AutoSize = $true
        $mailTextSalutationLabel.Text     = "Sehr geehrte Damen und Herren,"
        $mailTextSalutationLabel.Font     = $standradfont
        #endregion

        #region Mail-Text-Paket-Ankündigung
        $mailTextIntroductionLabel          = New-Object System.Windows.Forms.Label
        $mailTextIntroductionLabel.Location = New-Object System.Drawing.Point(10,30)
        $mailTextIntroductionLabel.AutoSize = $true
        $mailTextIntroductionLabel.Text     = 'das Team Rollout des ITZBunds wurde beauftragt, das Softwarepaket "'+$($mailSubjectPackageName.Text)+'" zu verteilen.'
        $mailTextIntroductionLabel.Font     = $standradfont

        $panelMailTextSalutation.Controls.Add($mailTextSalutationLabel)
        $panelMailTextSalutation.Controls.Add($mailTextIntroductionLabel)
        $panelMailText.Controls.Add($panelMailTextSalutation)
        #endregion

        #region Mail-Text-APCS
        $panelMailTextDevices             = New-Object System.Windows.Forms.Panel
        $panelMailTextDevices.Location    = New-Object System.Drawing.Point(0,95)
        $panelMailTextDevices.Size        = New-Object System.Drawing.Size(1002,30)
        $panelMailTextDevices.BorderStyle = "Fixed3D"

        $mailTextDeviceLabel          = New-Object System.Windows.Forms.Label
        $mailTextDeviceLabel.Location = New-Object System.Drawing.Point(10,7)
        $mailTextDeviceLabel.AutoSize = $true
        $mailTextDeviceLabel.Text     = "Welche APC's sind betroffen:"
        $mailTextDeviceLabel.Font     = $standradfont

        $mailTextAllDevicesComboBox          = New-Object System.Windows.Forms.CheckBox
        $mailTextAllDevicesComboBox.Location = New-Object System.Drawing.Point(530,7)
        $mailTextAllDevicesComboBox.AutoSize = $true
        $mailTextAllDevicesComboBox.Font     = $standradfont
        $mailTextAllDevicesComboBox.Text     = "ALLE"
        $mailTextAllDevicesComboBox.Checked  = $true
        
        $mailTextAllDevicesComboBox.add_CheckedChanged({
            if ($mailTextAllDevicesComboBox.Checked) {
                $mailTextDevicesComboBox.Checked = $false
            }
        })

        $mailTextDevicesComboBox          = New-Object System.Windows.Forms.CheckBox
        $mailTextDevicesComboBox.Location = New-Object System.Drawing.Point(600,7)
        $mailTextDevicesComboBox.AutoSize = $true
        $mailTextDevicesComboBox.Font     = $standradfont
        $mailTextDevicesComboBox.Text     = "Nur APC's mit einer installierter Vorversion"

        $mailTextDevicesComboBox.add_CheckedChanged({
            if ($mailTextDevicesComboBox.Checked) {
                $mailTextAllDevicesComboBox.Checked = $false
            }
        })

        $panelMailTextDevices.Controls.Add($mailTextDeviceLabel)
        $panelMailTextDevices.Controls.Add($mailTextAllDevicesComboBox)
        $panelMailTextDevices.Controls.Add($mailTextDevicesComboBox)
        $panelMailText.Controls.Add($panelMailTextDevices)
        #endregion
        
        #region Mail-Text-Umgebungen
        $panelMailTextEnvironment             = New-Object System.Windows.Forms.Panel
        $panelMailTextEnvironment.Location    = New-Object System.Drawing.Point(0,125)
        $panelMailTextEnvironment.Size        = New-Object System.Drawing.Size(1002,65)
        $panelMailTextEnvironment.BorderStyle = "Fixed3D"

        $mailTextEnvironmentLabel          = New-Object System.Windows.Forms.Label
        $mailTextEnvironmentLabel.Location = New-Object System.Drawing.Point(10,25)
        $mailTextEnvironmentLabel.AutoSize = $true
        $mailTextEnvironmentLabel.Text     = "Welche Umgebungen sind betroffen:"
        $mailTextEnvironmentLabel.Font     = $standradfont

        # Erstellen einer CheckedListBox
        $mailTextEnvironmentListBox              = New-Object System.Windows.Forms.CheckedListBox
        $mailTextEnvironmentListBox.Location     = New-Object System.Drawing.Point(530,5)
        $mailTextEnvironmentListBox.Size         = New-Object System.Drawing.Size(200,60)
        $mailTextEnvironmentListBox.Font         = $standradfont
        $mailTextEnvironmentListBox.CheckOnClick = $true

        $mailTextEnvironmentListBox.Add_SelectedValueChanged({
            $checked = 0
            
            For($i=1; $i -lt $mailTextEnvironmentListBox.Items.Count; $i++)
            {
                If( $mailTextEnvironmentListBox.GetItemCheckState($i) -eq "Checked")
                {
                    $checked++
                }   
            }

            If($checked -eq ($mailTextEnvironmentListBox.Items.Count - 1) )
            {
                 for ($i = 1; $i -lt $mailTextEnvironmentListBox.Items.Count; $i++)
                 {
                    $mailTextEnvironmentListBox.SetItemChecked($i, $false)
                 }
                 $mailTextEnvironmentListBox.SetItemChecked(0, $true)
            }
        })

        $mailTextEnvironmentListBox.Add_Click({
            $selectedItem = $mailTextEnvironmentListBox.SelectedItem

            # Überprüfen, ob Element 1 ausgewählt ist
            if ($selectedItem -eq "Alle Kunden") {
                # Wenn Element 1 ausgewählt ist und ein anderes Element geklickt wird, deaktivieren Sie Element 1
                if ($mailTextEnvironmentListBox.CheckedIndices.Contains(0) -and $mailTextEnvironmentListBox.SelectedIndex -ne 0)
                {
                    $mailTextEnvironmentListBox.SetItemChecked(0, $false)
                }
                # Wenn Element 1 deaktiviert ist und es ausgewählt wird, deaktivieren Sie alle anderen Elemente
                elseif (-not $mailTextEnvironmentListBox.CheckedIndices.Contains(0) -and $mailTextEnvironmentListBox.SelectedIndex -eq 0)
                {
                    for ($i = 1; $i -lt $mailTextEnvironmentListBox.Items.Count; $i++)
                    {
                        $mailTextEnvironmentListBox.SetItemChecked($i, $false)
                    }
                }
            }
            # Wenn ein anderes Element als Element 1 ausgewählt ist, deaktivieren Sie Element 1
            else {
                $mailTextEnvironmentListBox.SetItemChecked(0, $false)
            }
        })

         # Hinzufügen von Items in die CheckedListBox
        for ($i = 0; $i -lt $mailEnvironmentList.Count; $i++) {
            $mailTextEnvironmentListBox.Items.Add($mailEnvironmentList[$i]) | Out-Null
        }

        $mailTextEnvironmentListBox.SetItemChecked(0, $true)

        $panelMailTextEnvironment.Controls.Add($mailTextEnvironmentLabel)
        $panelMailTextEnvironment.Controls.Add($mailTextEnvironmentListBox)
        $panelMailText.Controls.Add($panelMailTextEnvironment)
        #endregion

        #region Mail-Text-Rollout-Datum
        $panelMailTextDate             = New-Object System.Windows.Forms.Panel
        $panelMailTextDate.Location    = New-Object System.Drawing.Point(0,190)
        $panelMailTextDate.Size        = New-Object System.Drawing.Size(1002,60)
        $panelMailTextDate.BorderStyle = "Fixed3D"

        $mailTextDateLabel          = New-Object System.Windows.Forms.Label
        $mailTextDateLabel.Location = New-Object System.Drawing.Point(10,5)
        $mailTextDateLabel.AutoSize = $true
        $mailTextDateLabel.Text     = "Start des Rollouts, ab dem:"
        $mailTextDateLabel.Font     = $standradfont

        $mailTextDateSelector          = New-Object System.Windows.Forms.DateTimePicker
        $mailTextDateSelector.Location = New-Object System.Drawing.Point(530, 5)
        $mailTextDateSelector.Width    = 200

        $mailTextDateDeadLineLabel          = New-Object System.Windows.Forms.Label
        $mailTextDateDeadLineLabel.Location = New-Object System.Drawing.Point(10,30)
        $mailTextDateDeadLineLabel.AutoSize = $true
        $mailTextDateDeadLineLabel.Text     = "Handelt es sich um eine Stichtagsumstellung:"
        $mailTextDateDeadLineLabel.Font     = $standradfont

        $mailTextDateDeadLineYesComboBox          = New-Object System.Windows.Forms.CheckBox
        $mailTextDateDeadLineYesComboBox.Location = New-Object System.Drawing.Point(530,30)
        $mailTextDateDeadLineYesComboBox.AutoSize = $true
        $mailTextDateDeadLineYesComboBox.Font     = $standradfont
        $mailTextDateDeadLineYesComboBox.Text     = "Ja"
        
        $mailTextDateDeadLineYesComboBox.add_CheckedChanged({
            if ($mailTextDateDeadLineYesComboBox.Checked) {
                $mailTextDateDeadLineNoComboBox.Checked = $false
            }
        })

        $mailTextDateDeadLineNoComboBox          = New-Object System.Windows.Forms.CheckBox
        $mailTextDateDeadLineNoComboBox.Location = New-Object System.Drawing.Point(600,30)
        $mailTextDateDeadLineNoComboBox.AutoSize = $true
        $mailTextDateDeadLineNoComboBox.Font     = $standradfont
        $mailTextDateDeadLineNoComboBox.Text     = "Nein"
        $mailTextDateDeadLineNoComboBox.Checked  = $true

        $mailTextDateDeadLineNoComboBox.add_CheckedChanged({
            if ($mailTextDateDeadLineNoComboBox.Checked) {
                $mailTextDateDeadLineYesComboBox.Checked = $false
            }
        })

        $panelMailTextDate.Controls.Add($mailTextDateLabel)
        $panelMailTextDate.Controls.Add($mailTextDateSelector)
        $panelMailTextDate.Controls.Add($mailTextDateDeadLineLabel)
        $panelMailTextDate.Controls.Add($mailTextDateDeadLineYesComboBox)
        $panelMailTextDate.Controls.Add($mailTextDateDeadLineNoComboBox)
        $panelMailText.Controls.Add($panelMailTextDate)
        #endregion

        #region Mail-Text-Sichheitsrelevantes-Update
        $panelMailTextSecurityPatch             = New-Object System.Windows.Forms.Panel
        $panelMailTextSecurityPatch.Location    = New-Object System.Drawing.Point(0,250)
        $panelMailTextSecurityPatch.Size        = New-Object System.Drawing.Size(1002,55)
        $panelMailTextSecurityPatch.BorderStyle = "Fixed3D"

        $mailTextSecurityPatchLabel          = New-Object System.Windows.Forms.Label
        $mailTextSecurityPatchLabel.Location = New-Object System.Drawing.Point(10,5)
        $mailTextSecurityPatchLabel.AutoSize = $true
        $mailTextSecurityPatchLabel.Text     = "Handelt es sich um ein sicherheitsrelevantes / zeitkritisches Softwarepaket / Update:"
        $mailTextSecurityPatchLabel.Font     = $standradfont

        $mailTextSecurityPatchYesComboBox          = New-Object System.Windows.Forms.CheckBox
        $mailTextSecurityPatchYesComboBox.Location = New-Object System.Drawing.Point(530,5)
        $mailTextSecurityPatchYesComboBox.AutoSize = $true
        $mailTextSecurityPatchYesComboBox.Font     = $standradfont
        $mailTextSecurityPatchYesComboBox.Text     = "Ja"
        
        $mailTextSecurityPatchYesComboBox.add_CheckedChanged({
            if ($mailTextSecurityPatchYesComboBox.Checked) {
                $mailTextSecurityPatchNoComboBox.Checked = $false
                $mailTextSecurityPatchYesLabel.Visible  = $true
            }
        })

        $mailTextSecurityPatchNoComboBox          = New-Object System.Windows.Forms.CheckBox
        $mailTextSecurityPatchNoComboBox.Location = New-Object System.Drawing.Point(600,5)
        $mailTextSecurityPatchNoComboBox.AutoSize = $true
        $mailTextSecurityPatchNoComboBox.Font     = $standradfont
        $mailTextSecurityPatchNoComboBox.Text     = "Nein"
        $mailTextSecurityPatchNoComboBox.Checked  = $true

        $mailTextSecurityPatchNoComboBox.add_CheckedChanged({
            if ($mailTextSecurityPatchNoComboBox.Checked) {
                $mailTextSecurityPatchYesComboBox.Checked = $false
                $mailTextSecurityPatchYesLabel.Visible  = $false
            }
        })

        $mailTextSecurityPatchYesLabel          = New-Object System.Windows.Forms.Label
        $mailTextSecurityPatchYesLabel.Location = New-Object System.Drawing.Point(10,30)
        $mailTextSecurityPatchYesLabel.AutoSize = $true
        $mailTextSecurityPatchYesLabel.Text     = "Wir bitten die Kurzfristigkeit zu entschuldigen, aber es handelt sich um ein sicherheitsrelevantes Paket."
        $mailTextSecurityPatchYesLabel.Font     = $standradfont
        $mailTextSecurityPatchYesLabel.Visible  = $false

        $panelMailTextSecurityPatch.Controls.Add($mailTextSecurityPatchLabel)
        $panelMailTextSecurityPatch.Controls.Add($mailTextSecurityPatchYesComboBox)
        $panelMailTextSecurityPatch.Controls.Add($mailTextSecurityPatchNoComboBox)
        $panelMailTextSecurityPatch.Controls.Add($mailTextSecurityPatchYesLabel)
        $panelMailText.Controls.Add($panelMailTextSecurityPatch)
        #endregion

        #region Mail-Text-Update-Mechanismus
        $panelMailTextSecurityPatch             = New-Object System.Windows.Forms.Panel
        $panelMailTextSecurityPatch.Location    = New-Object System.Drawing.Point(0,305)
        $panelMailTextSecurityPatch.Size        = New-Object System.Drawing.Size(1002,65)
        $panelMailTextSecurityPatch.BorderStyle = "Fixed3D"

        $mailTextPatchRoutineLabel          = New-Object System.Windows.Forms.Label
        $mailTextPatchRoutineLabel.Location = New-Object System.Drawing.Point(10,23)
        $mailTextPatchRoutineLabel.AutoSize = $true
        $mailTextPatchRoutineLabel.Text     = "Art der Installion / des Updates:"
        $mailTextPatchRoutineLabel.Font     = $standradfont
        
         # Erstellen einer CheckedListBox
        $mailTextPatchRoutineTypeListBox              = New-Object System.Windows.Forms.CheckedListBox
        $mailTextPatchRoutineTypeListBox.Location     = New-Object System.Drawing.Point(530,5)
        $mailTextPatchRoutineTypeListBox.Size         = New-Object System.Drawing.Size(350,60)
        $mailTextPatchRoutineTypeListBox.Font         = $standradfont
        $mailTextPatchRoutineTypeListBox.CheckOnClick = $true
        $mailTextPatchRoutineTypeListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::One

        <#
        $mailTextPatchRoutineTypeListBox.Add_SelectedValueChanged({
            $checked = 0
            
            For($i=1; $i -lt $mailTextPatchRoutineTypeListBox.Items.Count; $i++)
            {
                If( $mailTextPatchRoutineTypeListBox.GetItemCheckState($i) -eq "Checked")
                {
                    $checked++
                }   
            }

            If($checked -eq ($mailTextPatchRoutineTypeListBox.Items.Count - 1) )
            {
                 for ($i = 1; $i -lt $mailTextPatchRoutineTypeListBox.Items.Count; $i++)
                 {
                    $mailTextPatchRoutineTypeListBox.SetItemChecked($i, $false)
                 }
                 $mailTextPatchRoutineTypeListBox.SetItemChecked(0, $true)
            }
        })

        $mailTextPatchRoutineTypeListBox.Add_Click({
            $selectedItem = $mailTextPatchRoutineTypeListBox.SelectedItem

            # Überprüfen, ob Element 1 ausgewählt ist
            if ($selectedItem -eq "Alle Kunden") {
                # Wenn Element 1 ausgewählt ist und ein anderes Element geklickt wird, deaktivieren Sie Element 1
                if ($mailTextPatchRoutineTypeListBox.CheckedIndices.Contains(0) -and $mailTextPatchRoutineTypeListBox.SelectedIndex -ne 0)
                {
                    $mailTextPatchRoutineTypeListBox.SetItemChecked(0, $false)
                }
                # Wenn Element 1 deaktiviert ist und es ausgewählt wird, deaktivieren Sie alle anderen Elemente
                elseif (-not $mailTextPatchRoutineTypeListBox.CheckedIndices.Contains(0) -and $mailTextPatchRoutineTypeListBox.SelectedIndex -eq 0)
                {
                    for ($i = 1; $i -lt $mailTextPatchRoutineTypeListBox.Items.Count; $i++)
                    {
                        $mailTextPatchRoutineTypeListBox.SetItemChecked($i, $false)
                    }
                }
            }
            # Wenn ein anderes Element als Element 1 ausgewählt ist, deaktivieren Sie Element 1
            else {
                $mailTextPatchRoutineTypeListBox.SetItemChecked(0, $false)
            }
        })
        #>
         # Hinzufügen von Items in die CheckedListBox
        for ($i = 0; $i -lt $mailInstallTypeList.Count; $i++) {
            $mailTextPatchRoutineTypeListBox.Items.Add($mailInstallTypeList[$i]) | Out-Null
        }

        $mailTextPatchRoutineTypeListBox.SetItemChecked(0, $true)

        $panelMailTextSecurityPatch.Controls.Add($mailTextPatchRoutineLabel)
        $panelMailTextSecurityPatch.Controls.Add($mailTextPatchRoutineTypeListBox)
        $panelMailText.Controls.Add($panelMailTextSecurityPatch)
        #endregion

        #region Mail-Text-Reboot
        $panelMailTextReboot             = New-Object System.Windows.Forms.Panel
        $panelMailTextReboot.Location    = New-Object System.Drawing.Point(0,370)
        $panelMailTextReboot.Size        = New-Object System.Drawing.Size(1002,30)
        $panelMailTextReboot.BorderStyle = "Fixed3D"

        $mailTextRebootLabel          = New-Object System.Windows.Forms.Label
        $mailTextRebootLabel.Location = New-Object System.Drawing.Point(10,7)
        $mailTextRebootLabel.AutoSize = $true
        $mailTextRebootLabel.Text     = "Ist ein Neustart des APC’s notwendig:"
        $mailTextRebootLabel.Font     = $standradfont
        
        $mailTextRebootYesCheckBox          = New-Object System.Windows.Forms.CheckBox
        $mailTextRebootYesCheckBox.Location = New-Object System.Drawing.Point(530,5)
        $mailTextRebootYesCheckBox.AutoSize = $true
        $mailTextRebootYesCheckBox.Font     = $standradfont
        $mailTextRebootYesCheckBox.Text     = "Ja"
        
        $mailTextRebootYesCheckBox.add_CheckedChanged({
            if ($mailTextRebootYesCheckBox.Checked) {
                $mailTextRebootNoCheckBox.Checked = $false
            }
        })

        $mailTextRebootNoCheckBox          = New-Object System.Windows.Forms.CheckBox
        $mailTextRebootNoCheckBox.Location = New-Object System.Drawing.Point(600,5)
        $mailTextRebootNoCheckBox.AutoSize = $true
        $mailTextRebootNoCheckBox.Font     = $standradfont
        $mailTextRebootNoCheckBox.Text     = "Nein"
        $mailTextRebootNoCheckBox.Checked  = $true

        $mailTextRebootNoCheckBox.add_CheckedChanged({
            if ($mailTextRebootNoCheckBox.Checked) {
                $mailTextRebootYesCheckBox.Checked = $false
            }
        })
        
        $panelMailTextReboot.Controls.Add($mailTextRebootLabel)
        $panelMailTextReboot.Controls.Add($mailTextRebootYesCheckBox)
        $panelMailTextReboot.Controls.Add($mailTextRebootNoCheckBox)
        $panelMailText.Controls.Add($panelMailTextReboot)

        #endregion

        #region Mail-Text-Weitere-Informationen
        $panelMailTextOtherInformations             = New-Object System.Windows.Forms.Panel
        $panelMailTextOtherInformations.Location    = New-Object System.Drawing.Point(0,400)
        $panelMailTextOtherInformations.Size        = New-Object System.Drawing.Size(1002,80)
        $panelMailTextOtherInformations.BorderStyle = "Fixed3D"

        $mailTextOtherInformationsLabel          = New-Object System.Windows.Forms.Label
        $mailTextOtherInformationsLabel.Location = New-Object System.Drawing.Point(10,5)
        $mailTextOtherInformationsLabel.AutoSize = $true
        $mailTextOtherInformationsLabel.Text     = "Allgemeine Infomationen:"
        $mailTextOtherInformationsLabel.Font     = $standradfont
        
        $mailTextOtherInformationsRichTextBox          = New-Object System.Windows.Forms.RichTextBox
        $mailTextOtherInformationsRichTextBox.Location = New-Object System.Drawing.Point(10,25)
        $mailTextOtherInformationsRichTextBox.Size     = New-Object System.Drawing.Size(985,50)
        $mailTextOtherInformationsRichTextBox.Text     = "Keine / Prozessprüfung / Manuelle Installation per Softwarecenter startet / Tasksequenz / Abhängigkeiten /`nAuf Grund der Paketgröße, kann es zu länger Downloadzeiten im HomeOffe kommen / Rollout erfolgt in Gruppierungen (Mehrere Tage) / <Langer Text>"
        $mailTextOtherInformationsRichTextBox.Font     = $standradfont

        $mailTextOtherInformationsRichTextBox.Add_Click({
            $mailTextOtherInformationsRichTextBox.Text = ""
        })

        $mailTextOtherInformationsRichTextBox.Add_LostFocus({
            If( [String]::IsNullOrEmpty($mailTextOtherInformationsRichTextBox.Text) )
            {
                $mailTextOtherInformationsRichTextBox.Text = "Keine / Prozessprüfung / Manuelle Installation per Softwarecenter startet / Tasksequenz / Abhängigkeiten /`nAuf Grund der Paketgröße, kann es zu länger Downloadzeiten im HomeOffe kommen / Rollout erfolgt in Gruppierungen (Mehrere Tage) / <Langer Text>"
            }
        })

        $panelMailTextOtherInformations.Controls.Add($mailTextOtherInformationsLabel)
        $panelMailTextOtherInformations.Controls.Add($mailTextOtherInformationsRichTextBox)
        $panelMailText.Controls.Add($panelMailTextOtherInformations)
        #endregion

        #region Mail-Text-Inform-BWZ-INC
        $panelMailTextBWZInform             = New-Object System.Windows.Forms.Panel
        $panelMailTextBWZInform.Location    = New-Object System.Drawing.Point(0,480)
        $panelMailTextBWZInform.Size        = New-Object System.Drawing.Size(1002,95)
        $panelMailTextBWZInform.BorderStyle = "Fixed3D"

        $mailTextUserInfoLabel          = New-Object System.Windows.Forms.Label
        $mailTextUserInfoLabel.Location = New-Object System.Drawing.Point(10,5)
        $mailTextUserInfoLabel.AutoSize = $true
        $mailTextUserInfoLabel.Text     = "Bitte informieren Sie ggfs. Ihre Anwender/innen."
        $mailTextUserInfoLabel.Font     = $standradfont

        $mailTextBWZLabel          = New-Object System.Windows.Forms.Label
        $mailTextBWZLabel.Location = New-Object System.Drawing.Point(10,30)
        $mailTextBWZLabel.AutoSize = $true
        $mailTextBWZLabel.Text     = "@BWZ:`nWir gehen davon aus, dass Ihre Geräte mitversorgen werden sollen, ansonst bitte formlos bei mir melden."
        $mailTextBWZLabel.Font     = $standradfont

        $mailTextINCLabel          = New-Object System.Windows.Forms.Label
        $mailTextINCLabel.Location = New-Object System.Drawing.Point(10,70)
        $mailTextINCLabel.AutoSize = $true
        $mailTextINCLabel.Text     = "Im Fall von vereinzelten Fehlern, eröffnen Sie bitte wie gewohnt ein Störungsticket (Incident) über das Ticketsystem."
        $mailTextINCLabel.Font     = $standradfont

        $panelMailTextBWZInform.Controls.Add($mailTextUserInfoLabel)
        $panelMailTextBWZInform.Controls.Add($mailTextBWZLabel)
        $panelMailTextBWZInform.Controls.Add($mailTextINCLabel)
        $panelMailText.Controls.Add($panelMailTextBWZInform)

        #endregion

        #region Verantwortlicher
        $panelMailTextResponsible            = New-Object System.Windows.Forms.Panel
        $panelMailTextResponsible.Location    = New-Object System.Drawing.Point(0,575)
        $panelMailTextResponsible.Size        = New-Object System.Drawing.Size(1002,65)
        $panelMailTextResponsible.BorderStyle = "Fixed3D"

        $mailTextResponsibleLabel          = New-Object System.Windows.Forms.Label
        $mailTextResponsibleLabel.Location = New-Object System.Drawing.Point(10,25)
        $mailTextResponsibleLabel.AutoSize = $true
        $mailTextResponsibleLabel.Font     = $standradfont
        $mailTextResponsibleLabel.Text     = "Verantwortlich für den Rollout ist"

        # Erstellen einer CheckedListBox
        $mailTextResponsibleListBox              = New-Object System.Windows.Forms.CheckedListBox
        $mailTextResponsibleListBox.Location     = New-Object System.Drawing.Point(530,5)
        $mailTextResponsibleListBox.Size         = New-Object System.Drawing.Size(200,60)
        $mailTextResponsibleListBox.Font         = $standradfont
        $mailTextResponsibleListBox.CheckOnClick = $true

        $mailTextResponsibleListBox.Add_SelectedValueChanged({
            $checked = 0
            
            For($i=1; $i -lt $mailTextResponsibleListBox.Items.Count; $i++)
            {
                If( $mailTextResponsibleListBox.GetItemCheckState($i) -eq "Checked")
                {
                    $checked++
                }   
            }

            If($checked -eq ($mailTextResponsibleListBox.Items.Count - 1) )
            {
                 for ($i = 1; $i -lt $mailTextResponsibleListBox.Items.Count; $i++)
                 {
                    $mailTextResponsibleListBox.SetItemChecked($i, $false)
                 }
                 $mailTextResponsibleListBox.SetItemChecked(0, $true)
            }
        })

        $mailTextResponsibleListBox.Add_Click({
            $selectedItem = $mailTextEnvironmentListBox.SelectedItem

            # Überprüfen, ob Element 1 ausgewählt ist
            if ($selectedItem -eq "Alle Kunden") {
                # Wenn Element 1 ausgewählt ist und ein anderes Element geklickt wird, deaktivieren Sie Element 1
                if ($mailTextResponsibleListBox.CheckedIndices.Contains(0) -and $mailTextResponsibleListBox.SelectedIndex -ne 0)
                {
                    $mailTextEnvironmentListBox.SetItemChecked(0, $false)
                }
                # Wenn Element 1 deaktiviert ist und es ausgewählt wird, deaktivieren Sie alle anderen Elemente
                elseif (-not $mailTextResponsibleListBox.CheckedIndices.Contains(0) -and $mailTextResponsibleListBox.SelectedIndex -eq 0)
                {
                    for ($i = 1; $i -lt $mailTextResponsibleListBox.Items.Count; $i++)
                    {
                        $mailTextResponsibleListBox.SetItemChecked($i, $false)
                    }
                }
            }
            # Wenn ein anderes Element als Element 1 ausgewählt ist, deaktivieren Sie Element 1
            else {
                $mailTextResponsibleListBox.SetItemChecked(0, $false)
            }
        })

         # Hinzufügen von Items in die CheckedListBox
        for ($i = 0; $i -lt $mailResponsible.Count; $i++) {
            $mailTextResponsibleListBox.Items.Add($mailResponsible[$i]) | Out-Null
        }

        $mailTextResponsibleListBox.SetItemChecked(0, $true)

        $panelMailTextResponsible.Controls.Add($mailTextResponsibleLabel)
        $panelMailTextResponsible.Controls.Add($mailTextResponsibleListBox)
        $panelMailText.Controls.Add($panelMailTextResponsible)
        #endregion

        #region Mail-Text-Ende
        $panelMailTextEnd             = New-Object System.Windows.Forms.Panel
        $panelMailTextEnd.Location    = New-Object System.Drawing.Point(0,640)
        $panelMailTextEnd.Size        = New-Object System.Drawing.Size(1002,30)
        $panelMailTextEnd.BorderStyle = "Fixed3D"

        $mailTextEndLabel          = New-Object System.Windows.Forms.Label
        $mailTextEndLabel.Location = New-Object System.Drawing.Point(10,5)
        $mailTextEndLabel.AutoSize = $true
        $mailTextEndLabel.Text     = "Dies ist eine allgemeine Information, wir bitten von Einzel-E-Mail-Verkehr abzusehen!"
        $mailTextEndLabel.Font     = $standradfont

        $panelMailTextEnd.Controls.Add($mailTextEndLabel)
        $panelMailText.Controls.Add($panelMailTextEnd)
        #endregion

        #region Hinzufügen-von-Komponenten-zum-Hauptfenster
        # Komponenten zum Hauptfenster hinzufügen
        $mainWindow.Controls.Add($panelMailFrom)
        $mainWindow.Controls.Add($panelMailTo)
        $mainWindow.Controls.Add($panelSubject)
        $mainWindow.Controls.Add($panelMailCC)
        $mainWindow.Controls.Add($panelMailText)
        #endregion
        
        # Anzeigen des Hauptfensters
        $mainWindow.ShowDialog() | Out-Null
    }

    End
    {}
}

MainWindow