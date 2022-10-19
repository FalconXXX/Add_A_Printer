#===========================================================#
#           add a printer \ *<*/  (for Users)               #
#           11.05.2019 					    #
#	    by FalconXXX 		 		    #
#           Build & Tested on:       Windows 10   & 11	    # 
#	    https://github.com/FalconXXX/PowerShell.git     #
#	    add the path ;)                                 #
#===========================================================#




#MainWindow
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


$window = New-Object System.Windows.Forms.Form
$window.Width = 800
$window.Height = 400
$window.Text = "add a printer \ *<*/"
	#Text im Fenster
	$Label = New-Object System.Windows.Forms.Label
	$Label.Location = New-Object System.Drawing.Size(40,10)
	$Label.Text = "Text`nText`nBitte geben Sie den Druckernamen an:"
	$Label.AutoSize = $True
	$window.Controls.Add($Label)
	 
	 #input field
	  $windowTextBox = New-Object System.Windows.Forms.TextBox
	  $windowTextBox.Location = New-Object System.Drawing.Size(10,80)
	  $windowTextBox.Size = New-Object System.Drawing.Size(500,500)
	  $window.Controls.Add($windowTextBox)
	
	  #add Button
	  $addButton = New-Object System.Windows.Forms.Button
	  $addButton.Location = New-Object System.Drawing.Size(10,100)
	  $addButton.Size = New-Object System.Drawing.Size(80,30)
	  $addButton.Text = "Hinzufuegen"
	  $addButton.Add_Click({searchPrinter})	 
	  $window.Controls.Add($addButton)  #erzeug den Button

	#stop Button
	  $stopButton = New-Object System.Windows.Forms.Button
	  $stopButton.Location = New-Object System.Drawing.Size(100,100)
	  $stopButton.Size = New-Object System.Drawing.Size(80,30)
	  $stopButton.Text = "Abbrechen"
	  $stopButton.Add_Click({ $window.Dispose()})	
	  $window.Controls.Add($stopButton) #erzeug den Button
	  
	  #remove Button
	  $removeButton = New-Object System.Windows.Forms.Button
	  $removeButton.Location = New-Object System.Drawing.Size(430,100)
	  $removeButton.Size = New-Object System.Drawing.Size(80,30)
	  $removeButton.Text = "Entfernen"
	  $removeButton.Add_Click({deletePrinter})
	  $window.Controls.Add($removeButton) #erzeug den Button
	  
	  #check default printer
		$CHKStandard= New-Object System.Windows.Forms.CheckBox
		$CHKStandard.Location = New-Object System.Drawing.Size(200,100)
		$CHKStandard.Size = New-Object System.Drawing.Size(120,30)
		$CHKStandard.Text = "Standarddrucker"
		$CHKStandard.Checked = $false
		$window.Controls.Add($CHKStandard)
			  
	  
	  
	  #function searchPrinter
	
		function searchPrinter
		{
			$TERMSRV = $env:computername
			if($TERMSRV.StartsWith("TERMSRV"))
			{
				Read-Host -Prompt "Sie befinden sich auf einem Terminalserver `nbitte ein Ticket oeffen"
				\\Path\addAPrinter3.bat
			}
			else
			{
				if(Test-Path C:\Path\Lexmark)
					{
						$p = $windowTextBox.Text 
						$n = $p.ToUpper()
						$temp
						$printersAvailable
						$tempPortName = 0
						if ($n.length -eq 8)
						{
							if($n.StartsWith("PRSC") -or $n.StartsWith("PRRX") -or $n.StartsWith("PRRG"))
							{
								$printerobjects = Import-CSV "\\Path\drucker11052019.csv"
								
								for ($i=0; $i -lt $printerobjects.length; $i++)
								{
										
										if($n -eq $printerobjects[$i].Name -and $n.length -eq 8)
											{
												if( $printerobjects[$i].DriverName -eq "TSC TC200" -or $printerobjects[$i].DriverName -eq "TSC TTP-2410MT")
												{
													$Ausgabe1.text =  "Bitte ein Ticket oeffen, hier handelt es sich um einen Etikettendrucker/Plotter "
													\\Path\addAPrinter3.bat
													
												}
												else
												{
													$printersAvailable = $n
													$tempPortName = $printerobjects[$i].PortName
												}
											}
											
								}
									if($printersAvailable -eq  $n)
										{
											Remove-PrinterPort -Name "$tempPortName"				-ErrorAction SilentlyContinue
											Remove-Printer -Name "$printersAvailable"				-ErrorAction SilentlyContinue
											$Ausgabe1.text = "      *********Einen kurzen Augenblick der Drucker: $printersAvailable wird hinzugefuegt *********"
											 
											Add-PrinterPort -Name "$tempPortName" -PrinterHostAddress "$tempPortName"		-ErrorAction SilentlyContinue
											Add-Printer -Name "$printersAvailable" -DriverName "Lexmark Universal v2 XL (2.8.0.0)" -PortName "$tempPortName"		-ErrorAction SilentlyContinue
											
											if($CHKStandard.Checked)
											{
												(New-Object -ComObject WScript.Network).SetDefaultPrinter("$printersAvailable")
											}
											$Ausgabe1.text = "Drucker: $printersAvailable wurde hinzugefuegt " 
											
										
											}
										else
										{
										
										$Ausgabe1.text = "Drucker nicht vorhanden, wenn es sich um einen Etikettendrucker/Plotter handelt `nhierfuer bitte ein Ticket erstellen "
										}
												
							}
							else
							{
								$Ausgabe1.text = "Drucker nicht vorhanden, Namen beginnen mit PRSC.. /PRRX.. / PRRG.."
								
							}
						}
						else
						{
								$Ausgabe1.text = "Drucker nicht gefunden oder falsche Eingabe: z.B.: PRRX0101"
								
						}
					}
					else
					{
							$Ausgabe1.text =  "Bitte ein Ticket oeffen, `n auf Ihrem Geraet sind keine Druckertreiber hinterlegt "
							\\Path\addAPrinter3.bat
					}
			}		
		}
	
	        
 # END function searchPrinter 
 # function deletePrinter
		function deletePrinter
		{
				Remove-Item C:Path\DruckerLokal.csv
				$p = $windowTextBox.Text 
				$n = $p.ToUpper()
				$tempPortName
				$printersAvailable
				if ($n.length -eq 8)
				{
						if($n.StartsWith("PRSC") -or $n.StartsWith("PRRX") -or $n.StartsWith("PRRG"))
						{
								$DeviceName = $env:computername
								Get-WmiObject -Class Win32_Printer -ComputerName $DeviceName | select Name, PortName | Export-Csv -Path C:Path\DruckerLokal.csv -NoClobber -Delimiter ","
								$printerobjects = Import-CSV "C:Path\DruckerLokal.csv"
							
								if ($printerobjects.Name -contains $n) 
									{ 
										$Ausgabe1.text  ="      *********Einen kurzen Augenblick der Drucker: $printersAvailable wird entfernt *********"
										Remove-Printer -Name "$n"								-ErrorAction SilentlyContinue
										Remove-Item C:Path\DruckerLokal.csv
										$Ausgabe1.text = "Drucker: $n wurde entfernt " 
										
									}
								else
									{
										Remove-Item C:Path\DruckerLokal.csv
										$Ausgabe1.text = "Drucker nicht vorhanden,  "
										
									}
						}
						else
						{
							$Ausgabe1.text = "Drucker nicht vorhanden, Namen beginnen mit PRSC.. /PRRX.. / PRRG.."	
						}
						
				}
				else
				{
					$Ausgabe1.text = "Drucker nicht gefunden oder falsche Eingabe: z.B.: PRRX0101"				
				}
		}
	
		
 
 # END function deletePrinter
 

		#output field
		$Ausgabe1 = New-Object System.Windows.Forms.Textbox
		$Ausgabe1.Location = New-Object System.Drawing.Size(00,160)
		$Ausgabe1.Size = New-Object System.Drawing.Size(900,250)
		$Ausgabe1.MultiLine = $true
		$Ausgabe1.ReadOnly = $true
		$Ausgabe1.AutoSize = $true
		$Ausgabe1.Font = "Lucida Console, 10 pt"
		
		$Ausgabe1.WordWrap = $false
		$window.Controls.Add($Ausgabe1)
		
		#add a pic
		$Bild1 = New-Object System.Windows.Forms.PictureBox
		$Bild1.Location = New-Object System.Drawing.Size(550,10)
		$Bild1.ImageLocation = "\\Path\printer.PNG"
		$Bild1.SizeMode = "AutoSize"
		$window.Controls.Add($Bild1)

	  
[void]$window.ShowDialog()
