#----------------------------------------------
#region Application Functions
#----------------------------------------------

#endregion Application Functions

#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Show-365-Toolkit_psf {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Define SAPIEN Types
	#----------------------------------------------
	try{
		[FolderBrowserModernDialog] | Out-Null
	}
	catch
	{
		Add-Type -ReferencedAssemblies ('System.Windows.Forms') -TypeDefinition  @" 
		using System;
		using System.Windows.Forms;
		using System.Reflection;

        namespace SAPIENTypes
        {
		    public class FolderBrowserModernDialog : System.Windows.Forms.CommonDialog
            {
                private System.Windows.Forms.OpenFileDialog fileDialog;
                public FolderBrowserModernDialog()
                {
                    fileDialog = new System.Windows.Forms.OpenFileDialog();
                    fileDialog.Filter = "Folders|\n";
                    fileDialog.AddExtension = false;
                    fileDialog.CheckFileExists = false;
                    fileDialog.DereferenceLinks = true;
                    fileDialog.Multiselect = false;
                    fileDialog.Title = "Select a folder";
                }

                public string Title
                {
                    get { return fileDialog.Title; }
                    set { fileDialog.Title = value; }
                }

                public string InitialDirectory
                {
                    get { return fileDialog.InitialDirectory; }
                    set { fileDialog.InitialDirectory = value; }
                }
                
                public string SelectedPath
                {
                    get { return fileDialog.FileName; }
                    set { fileDialog.FileName = value; }
                }

                object InvokeMethod(Type type, object obj, string method, object[] parameters)
                {
                    MethodInfo methInfo = type.GetMethod(method, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    return methInfo.Invoke(obj, parameters);
                }

                bool ShowOriginalBrowserDialog(IntPtr hwndOwner)
                {
                    using(FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                    {
                        folderBrowserDialog.Description = this.Title;
                        folderBrowserDialog.SelectedPath = !string.IsNullOrEmpty(this.SelectedPath) ? this.SelectedPath : this.InitialDirectory;
                        folderBrowserDialog.ShowNewFolderButton = false;
                        if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                        {
                            fileDialog.FileName = folderBrowserDialog.SelectedPath;
                            return true;
                        }
                        return false;
                    }
                }

                protected override bool RunDialog(IntPtr hwndOwner)
                {
                    if (Environment.OSVersion.Version.Major >= 6)
                    {      
                        try
                        {
                            bool flag = false;
                            System.Reflection.Assembly assembly = Assembly.Load("System.Windows.Forms, Version = 4.0.0.0, Culture = neutral, PublicKeyToken = b77a5c561934e089");
                            Type typeIFileDialog = assembly.GetType("System.Windows.Forms.FileDialogNative").GetNestedType("IFileDialog", BindingFlags.NonPublic);
                            uint num = 0;
                            object dialog = InvokeMethod(fileDialog.GetType(), fileDialog, "CreateVistaDialog", null);
                            InvokeMethod(fileDialog.GetType(), fileDialog, "OnBeforeVistaDialog", new object[] { dialog });
                            uint options = (uint)InvokeMethod(typeof(System.Windows.Forms.FileDialog), fileDialog, "GetOptions", null) | (uint)0x20;
                            InvokeMethod(typeIFileDialog, dialog, "SetOptions", new object[] { options });
                            Type vistaDialogEventsType = assembly.GetType("System.Windows.Forms.FileDialog").GetNestedType("VistaDialogEvents", BindingFlags.NonPublic);
                            object pfde = Activator.CreateInstance(vistaDialogEventsType, fileDialog);
                            object[] parameters = new object[] { pfde, num };
                            InvokeMethod(typeIFileDialog, dialog, "Advise", parameters);
                            num = (uint)parameters[1];
                            try
                            {
                                int num2 = (int)InvokeMethod(typeIFileDialog, dialog, "Show", new object[] { hwndOwner });
                                flag = 0 == num2;
                            }
                            finally
                            {
                                InvokeMethod(typeIFileDialog, dialog, "Unadvise", new object[] { num });
                                GC.KeepAlive(pfde);
                            }
                            return flag;
                        }
                        catch
                        {
                            return ShowOriginalBrowserDialog(hwndOwner);
                        }
                    }
                    else
                        return ShowOriginalBrowserDialog(hwndOwner);
                }

                public override void Reset()
                {
                    fileDialog.Reset();
                }
            }
       }
"@ -IgnoreWarnings | Out-Null
	}
	#endregion Define SAPIEN Types

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form_365toolkit = New-Object 'System.Windows.Forms.Form'
	$tabcontrol1 = New-Object 'System.Windows.Forms.TabControl'
	$tabpage_Suite = New-Object 'System.Windows.Forms.TabPage'
	$listview_apps = New-Object 'System.Windows.Forms.ListView'
	$combobox_suite = New-Object 'System.Windows.Forms.ComboBox'
	$labelIncludedApps = New-Object 'System.Windows.Forms.Label'
	$labelOfficeSuite = New-Object 'System.Windows.Forms.Label'
	$groupbox1 = New-Object 'System.Windows.Forms.GroupBox'
	$radiobutton_64bit = New-Object 'System.Windows.Forms.RadioButton'
	$radiobutton_32bit = New-Object 'System.Windows.Forms.RadioButton'
	$tabpage2 = New-Object 'System.Windows.Forms.TabPage'
	$groupbox_Activation = New-Object 'System.Windows.Forms.GroupBox'
	$linklabelLearnMoreAboutDevice = New-Object 'System.Windows.Forms.LinkLabel'
	$radiobutton_deviceBased = New-Object 'System.Windows.Forms.RadioButton'
	$textbox_TokenPath = New-Object 'System.Windows.Forms.TextBox'
	$checkbox_SCLCacheOverride = New-Object 'System.Windows.Forms.CheckBox'
	$linklabelLearnMoreAboutShareC = New-Object 'System.Windows.Forms.LinkLabel'
	$linklabelLearnMoreAboutUserBa = New-Object 'System.Windows.Forms.LinkLabel'
	$radiobuttonSharedComputer = New-Object 'System.Windows.Forms.RadioButton'
	$radiobuttonUserBased = New-Object 'System.Windows.Forms.RadioButton'
	$richtextbox_eula = New-Object 'System.Windows.Forms.RichTextBox'
	$checkbox_AcceptEULA = New-Object 'System.Windows.Forms.CheckBox'
	$tabpage1 = New-Object 'System.Windows.Forms.TabPage'
	$splitcontainer1 = New-Object 'System.Windows.Forms.SplitContainer'
	$menustrip_main = New-Object 'System.Windows.Forms.MenuStrip'
	$fileToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$viewToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$button_Install = New-Object 'System.Windows.Forms.Button'
	$button_JustDownload = New-Object 'System.Windows.Forms.Button'
	$uIFontSettingsToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$fontdialog_main = New-Object 'System.Windows.Forms.FontDialog'
	$tooltip_main = New-Object 'System.Windows.Forms.ToolTip'
	$folderbrowsermoderndialog_main = New-Object 'SAPIENTypes.FolderBrowserModernDialog'
	$richtextbox_ProcessInfo = New-Object 'System.Windows.Forms.RichTextBox'
	$exitToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$aboutToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$checkForUpdatesToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$exportCurrentXMLToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$installXMLToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$downloadXMLToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$savefiledialog_main = New-Object 'System.Windows.Forms.SaveFileDialog'
	$columnheader1 = New-Object 'System.Windows.Forms.ColumnHeader'
	$columnheader2 = New-Object 'System.Windows.Forms.ColumnHeader'
	$timer_JobTracker = New-Object 'System.Windows.Forms.Timer'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	#region Functions
	function Download-ToolkitExec
	{
		param
		(
			[parameter()]
			[String]$Path = "C:\Windows\Temp\O365-Toolkit"
		)
		
		$URL = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_15028-20242.exe"
		
		if (!(Test-Path $Path))
		{
			$BasePath = $Path | Split-Path
			$FolderName = $Path | Split-Path -Leaf
			
			$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Provided path doesnt exist, creating it now...")
			New-Item -Path $BasePath -ItemType Directory -Name $FolderName -Force | Out-Null
		}
		
		# download the toolkit
		try
		{
			$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Begining download of Office 365 toolkit...`n")
			Invoke-WebRequest -Uri $URL -OutFile "$path\officedeploymenttool_15028-20242.exe"
		}
		Catch
		{
			# Powershell 3 and lower dl method here	
		}
		
		# extract the files from the toolkit
		try
		{
			$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Extracting Office 365 toolkit...`n")
			$process = Start-Process "$path\officedeploymenttool_15028-20242.exe" -ArgumentList "/extract:`"$Path`" /quiet" -PassThru -Wait
			$TimeElapsed = New-TimeSpan -Start $process.StartTime -End $process.ExitTime
			$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Download complete`n")
		}
		catch
		{
			
		}
		
		# return the full path to the toolkit exec
		Return "$Path\Setup.exe"	
	}
	function Generate-DownloadOfficeXML
	{
		param
		(
			[parameter(HelpMessage = "Path downloaded files will be saved to")]
			[string]$SourcePath = "C:\Windows\Temp\O365-Toolkit"
		)
		# Source path
		$SuiteOptions = @"
  <Add SourcePath="$SourcePath"
"@
		# Architecture
		if ($radiobutton_32bit.Checked)
		{
			$SuiteOptions += @"
 OfficeClientEdition="32"
"@
		}
		else
		{
			$SuiteOptions += @"
 OfficeClientEdition="64"
"@
		}
		
		# Update Channel
		$SuiteOptions += @"
 Channel="Current"
"@
		
		# Accept EULA
		if ($checkbox_AcceptEULA.Checked)
		{
			$SuiteOptions += @"
 AcceptEULA="TRUE">
"@
		}
		else
		{
			$SuiteOptions += @"
 AcceptEULA="FALSE">
"@
		}
		
		# Product Suite
		switch ($suite)
		{
			"Office 365 Apps for Enterprise" {
				$ProductID = @"
    <Product ID="O365ProPlusRetail">
"@
			}
			"Office 365 Apps for Business" {
				$ProductID = @"
    <Product ID="O365BusinessRetail">
"@
			}
			"Office LTSC Professional Plus 2021 - Volume License" {
				#<code>
			}
			"Office LTSC Professional Plus 2021 (SPLA) - Volume License" {
				#<code>
			}
			"Office LTSC Standard 2021 - Volume License"{
				
			}
			"Office LTSC Standard 2021 (SPLA) - Volume License"{
				
			}
			"Office Professional Plus 2019 - Volume License"{
				
			}
			"Office Standard 2019 - Volume License"{
				
			}
		}
		
		# Language
		$LanguageID = @"
      <Language ID="en-us" />
"@
		
		if ($checkbox_AcceptEULA.Checked)
		{
			$EulaInfo = @"
<Updates Enabled="TRUE" Channel="Current" />
<Display Level="None" AcceptEULA="TRUE" />
"@
		}
		else
		{
			$EulaInfo = @"
<Updates Enabled="TRUE" Channel="Current" />
<Display Level="None" AcceptEULA="FALSE" />
"@
		}
		
		# Construct XML
		$CompletedXML = @"
<Configuration>
$SuiteOptions
$productID
$LanguageID
    </Product>
  </Add> 
$EulaInfo
</Configuration>
"@
		Return $CompletedXML
	}
	function Generate-InstallOfficeXML
	{
		if ($DLPath)
		{
			# Office installation files are already downloaded, use that as source
			
			# Arch
			If ($radiobutton_32bit.Checked)
			{
				$ProductConfig = @"
  <Add OfficeClientEdition="32"
"@
			}
			else
			{
				$ProductConfig = @"
  <Add OfficeClientEdition="64"
"@
			}
			
			# Update Channel
			$ProductConfig += @"
 Channel="Current"
"@
			
			# Source Path
			$ProductConfig += @"
 SourcePath="$DLPath"
"@
			
			# CDN Fallback
			$ProductConfig += @"
 AllowCdnFallback="TRUE">
"@
			# Product Suite
			switch ($suite)
			{
				"Office 365 Apps for Enterprise" {
					$ProductID = @"
    <Product ID="O365ProPlusRetail">
"@
				}
				"Office 365 Apps for Business" {
					$ProductID = @"
    <Product ID="O365BusinessRetail">
"@
				}
				"Office LTSC Professional Plus 2021 - Volume License" {
					#<code>
				}
				"Office LTSC Professional Plus 2021 (SPLA) - Volume License" {
					#<code>
				}
				"Office LTSC Standard 2021 - Volume License"{
					
				}
				"Office LTSC Standard 2021 (SPLA) - Volume License"{
					
				}
				"Office Professional Plus 2019 - Volume License"{
					
				}
				"Office Standard 2019 - Volume License"{
					
				}
			}
			
			# Language ID
			$LanguageInfo = @"
      <Language ID="en-us" />
"@
			
			forEach ($item in $listview_apps.Items)
			{
				if ($item.Checked -eq $false)
				{
					switch ($item.text)
					{
						"Access" {
							$LanguageInfo += "`n      <ExcludeApp ID=`"Access`" />"
						}
						"OneDrive (Groove)" {
							$LanguageInfo += "`n      <ExcludeApp ID=`"Groove`" />"
						}
						"OneDrive Desktop" {
							$LanguageInfo += "`n      <ExcludeApp ID=`"OneDrive`" />"
						}
						"Outlook" {
							$LanguageInfo += "`n      <ExcludeApp ID=`"Outlook`" />"
						}
						"Publisher"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Publisher`" />"
						}
						"Word"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Word`" />"
						}
						"Excel"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Excel`" />"
						}
						"Skype for Business"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Lync`" />"
						}
						"OneNote"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"OneNote`" />"
						}
						"PowerPoint"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"PowerPoint`" />"
						}
						"Teams"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Teams`" />"
						}
					}
				}
				
			}
			
			if ($radiobuttonUserBased.Checked)
			{
				$LicenseInfo = @"
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
"@
			}
			else
			{
				$LicenseInfo = @"
  <Property Name="SharedComputerLicensing" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
"@
				
				if ($checkbox_SCLCacheOverride.checked)
				{
					$LicenseInfo += @"
`n  <Property Name="SCLCacheOverride" Value="1" />
  <Property Name="SCLCacheOverrideDirectory" Value="$($textbox_TokenPath.Text)" />
"@
				}
				else
				{
					$LicenseInfo += @"
`n  <Property Name="SCLCacheOverride" Value="0" />
"@
				}
			}
			
			# EULA
			if ($checkbox_AcceptEULA.Checked)
			{
				$LicenseInfo += @"
`n  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
"@
			}
			else
			{
				$LicenseInfo += @"
`n  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="FALSE" />
"@
			}
			
			# Complete the xml
			$CompletedXML = @"
<Configuration>
$ProductConfig
$productID
$LanguageInfo
    </Product>
  </Add>
$LicenseInfo
</Configuration>
"@
			Return $CompletedXML
		}
		else
		{
			# Arch
			If ($radiobutton_32bit.Checked)
			{
				$ProductConfig = @"
  <Add OfficeClientEdition="32"
"@
			}
			else
			{
				$ProductConfig = @"
  <Add OfficeClientEdition="64"
"@
			}
			
			# Update Channel
			$ProductConfig += @"
 Channel="Current"
"@
			
			# CDN Fallback
			$ProductConfig += @"
 AllowCdnFallback="TRUE">
"@
			# Product Suite
			switch ($suite)
			{
				"Office 365 Apps for Enterprise" {
					$ProductID = @"
    <Product ID="O365ProPlusRetail">
"@
				}
				"Office 365 Apps for Business" {
					$ProductID = @"
    <Product ID="O365BusinessRetail">
"@
				}
				"Office LTSC Professional Plus 2021 - Volume License" {
					#<code>
				}
				"Office LTSC Professional Plus 2021 (SPLA) - Volume License" {
					#<code>
				}
				"Office LTSC Standard 2021 - Volume License"{
					
				}
				"Office LTSC Standard 2021 (SPLA) - Volume License"{
					
				}
				"Office Professional Plus 2019 - Volume License"{
					
				}
				"Office Standard 2019 - Volume License"{
					
				}
			}
			
			# Language ID
			$LanguageInfo = @"
      <Language ID="en-us" />
"@
			
			forEach ($item in $listview_apps.Items)
			{
				if ($item.Checked -eq $false)
				{
					switch ($item.text)
					{
						"Access" {
							$LanguageInfo += "`n      <ExcludeApp ID=`"Access`" />"
						}
						"OneDrive (Groove)" {
							$LanguageInfo += "`n      <ExcludeApp ID=`"Groove`" />"
						}
						"OneDrive Desktop" {
							$LanguageInfo += "`n      <ExcludeApp ID=`"OneDrive`" />"
						}
						"Outlook" {
							$LanguageInfo += "`n      <ExcludeApp ID=`"Outlook`" />"
						}
						"Publisher"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Publisher`" />"
						}
						"Word"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Word`" />"
						}
						"Excel"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Excel`" />"
						}
						"Skype for Business"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Lync`" />"
						}
						"OneNote"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"OneNote`" />"
						}
						"PowerPoint"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"PowerPoint`" />"
						}
						"Teams"{
							$LanguageInfo += "`n      <ExcludeApp ID=`"Teams`" />"
						}
					}
				}
				
			}
			
			# Licensing
			if ($radiobuttonUserBased.Checked)
			{
				$LicenseInfo = @"
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
"@
			}
			else
			{
				$LicenseInfo = @"
  <Property Name="SharedComputerLicensing" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
"@
				
				if ($checkbox_SCLCacheOverride.checked)
				{
					$LicenseInfo += @"
`n  <Property Name="SCLCacheOverride" Value="1" />
  <Property Name="SCLCacheOverrideDirectory" Value="$($textbox_TokenPath.Text)" />
"@
				}
				else
				{
					$LicenseInfo += @"
`n  <Property Name="SCLCacheOverride" Value="0" />
"@
				}
			}
			
			# EULA
			if ($checkbox_AcceptEULA.Checked)
			{
				$LicenseInfo += @"
`n  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
"@
			}
			else
			{
				$LicenseInfo += @"
`n  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="FALSE" />
"@
			}
			
			# Complete the xml
			$CompletedXML = @"
<Configuration>
$ProductConfig
$productID
$LanguageInfo
    </Product>
  </Add>
$LicenseInfo
</Configuration>
"@
			Return $CompletedXML
		}
	}
	function Check-ApplicationUpdates
	{
		param
		(
			[parameter(Mandatory = $true)]
			[string]$URI,
			[parameter(Mandatory = $true)]
			$AppVersion,
			[parameter(Mandatory = $true)]
			[String]$AppName,
			[parameter()]
			[Switch]$Startup = $False
		)
		
		function Show-AppUpdaterPrompt
		{
			param
			(
				[parameter(Mandatory = $true)]
				[String]$AppName,
				[parameter(Mandatory = $true)]
				[String]$ReleaseNotes,
				[parameter(Mandatory = $true)]
				[String]$ReleaseURL
			)
			
			#----------------------------------------------
			#region Import the Assemblies
			#----------------------------------------------
			[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
			[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
			#endregion Import Assemblies
			
			#----------------------------------------------
			#region Generated Form Objects
			#----------------------------------------------
			[System.Windows.Forms.Application]::EnableVisualStyles()
			$form_Updates = New-Object 'System.Windows.Forms.Form'
			$buttonDismiss = New-Object 'System.Windows.Forms.Button'
			$linklabel_latestRelease = New-Object 'System.Windows.Forms.LinkLabel'
			$label_Title = New-Object 'System.Windows.Forms.Label'
			$label_ReleaseNotes = New-Object 'System.Windows.Forms.Label'
			$richtextbox1 = New-Object 'System.Windows.Forms.RichTextBox'
			$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
			#endregion Generated Form Objects
			
			#----------------------------------------------
			# User Generated Script
			#----------------------------------------------
			
			$form_Updates_Load = {
				$label_Title.Text = "$AppName application update found"
				$richtextbox1.Text = $ReleaseNotes
			}
			
			$linklabel_latestRelease_LinkClicked = [System.Windows.Forms.LinkLabelLinkClickedEventHandler]{
				#Event Argument: $_ = [System.Windows.Forms.LinkLabelLinkClickedEventArgs]
				Start-Process $ReleaseURL
			}
			
			$buttonDismiss_Click = {
				$form_Updates.Close()
			}
			# --End User Generated Script--
			#----------------------------------------------
			#region Generated Events
			#----------------------------------------------
			
			$Form_StateCorrection_Load =
			{
				#Correct the initial state of the form to prevent the .Net maximized form issue
				$form_Updates.WindowState = $InitialFormWindowState
			}
			
			$Form_Cleanup_FormClosed =
			{
				#Remove all event handlers from the controls
				try
				{
					$buttonDismiss.remove_Click($buttonDismiss_Click)
					$linklabel_latestRelease.remove_LinkClicked($linklabel_latestRelease_LinkClicked)
					$form_Updates.remove_Load($form_Updates_Load)
					$form_Updates.remove_Load($Form_StateCorrection_Load)
					$form_Updates.remove_FormClosed($Form_Cleanup_FormClosed)
				}
				catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
			}
			#endregion Generated Events
			
			#----------------------------------------------
			#region Generated Form Code
			#----------------------------------------------
			$form_Updates.SuspendLayout()
			#
			# form_Updates
			#
			$form_Updates.Controls.Add($buttonDismiss)
			$form_Updates.Controls.Add($linklabel_latestRelease)
			$form_Updates.Controls.Add($label_Title)
			$form_Updates.Controls.Add($label_ReleaseNotes)
			$form_Updates.Controls.Add($richtextbox1)
			$form_Updates.AutoScaleDimensions = New-Object System.Drawing.SizeF(10, 20)
			$form_Updates.AutoScaleMode = 'Font'
			$form_Updates.ClientSize = New-Object System.Drawing.Size(706, 331)
			#region Binary Data
			$Formatter_binaryFomatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
			$System_IO_MemoryStream = New-Object System.IO.MemoryStream ( ,[byte[]][System.Convert]::FromBase64String('
AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj00LjAuMC4wLCBD
dWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABNTeXN0
ZW0uRHJhd2luZy5JY29uAgAAAAhJY29uRGF0YQhJY29uU2l6ZQcEAhNTeXN0ZW0uRHJhd2luZy5T
aXplAgAAAAIAAAAJAwAAAAX8////E1N5c3RlbS5EcmF3aW5nLlNpemUCAAAABXdpZHRoBmhlaWdo
dAAACAgCAAAAMAAAADAAAAAPAwAAAL4lAAACAAABAAEAMDAAAAEAIACoJQAAFgAAACgAAAAwAAAA
YAAAAAEAIAAAAAAAACQAAMMOAADDDgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAIAAAA2AAAAeAAAAKsAAADSAAAA7QAAAPkAAAD5AAAA7QAAANEAAACr
AAAAdwAAADUAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbAAAAhwAAAOIAAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAADiAAAAhgAAABsAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEwAAAJUAAAD4AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAPgAAACTAAAAEgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AABXAAAA6wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA6gAAAFUAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAABQAAAJcAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAACUAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKAAAAtQAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAAswAAAAoAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAACvAAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAALQA
AAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAJgAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAACUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWAAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAAVQAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAASAAAA6wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA7AAAABQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAACXAAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAJMAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABwAAAD5AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
APgAAAAbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIcAAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAACGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAgAAAOIAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAADhAAAAAgAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOAAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAANQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAegAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAAdwAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAArgAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
qwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1AAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA0gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
7QAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA6wAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAA+QAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA+QAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAH8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB3AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAA7QAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAB/AAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAHcAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA6wAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1AAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA0gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArgAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAH8AAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB3AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAAqwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAegAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAA
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAAeAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOQAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAgAAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAANgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAA
AOMAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAH8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB3AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAADkAAAAAwAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAIgAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAACAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAACH
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABwAAAD5AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAAfwAA
AAAAAAAAAAAAdwAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAPkAAAAeAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAACZAAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAIAAAACAAAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAJUAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATAAAA6wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA7QAA
ABQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
WwAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAAVwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJsAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAACXAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAC2AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAALsAAAAG
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAKAAAAtgAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAAtQAAAAoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAJsAAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA
AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAACYAAAABQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAABaAAAA7QAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA7AAAAFgAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFQAAAJkAAAD5AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAPkA
AACXAAAAFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAdAAAAiAAAAOQAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAADjAAAAhwAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAA5AAAAegAAAK0AAADVAAAA
7gAAAPoAAAD6AAAA7gAAANQAAACtAAAAegAAADgAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAD///////8AAP///////wAA////////AAD///////8AAP//gAH//wAA//4A
AH//AAD/+AAAH/8AAP/wAAAP/wAA/8AAAAP/AAD/gAAAAf8AAP8AAAAA/wAA/wAAAAD/AAD+AAAA
AH8AAPwAAAAAPwAA/AAAAAA/AAD4AAAAAB8AAPgAD/AAHwAA8AAP8AAPAADwAA/wAA8AAPAAD/AA
DwAA8AAP8AAPAADwAA/wAA8AAPAAD/AADwAA8AAP8AAPAADwAf//gA8AAPAA//8ADwAA8AB//gAP
AADwAD/8AA8AAPAAH/gADwAA8AAP8AAPAADwAAfgAA8AAPgAA8AAHwAA+AABgAAfAAD8AAAAAD8A
APwAAAAAPwAA/gAAAAB/AAD/AAAAAP8AAP8AAAAA/wAA/4AAAAH/AAD/wAAAA/8AAP/wAAAP/wAA
//gAAB//AAD//gAAf/8AAP//gAH//wAA////////AAD///////8AAP///////wAA////////AAAL'))
			#endregion
			$form_Updates.Icon = $Formatter_binaryFomatter.Deserialize($System_IO_MemoryStream)
			$Formatter_binaryFomatter = $null
			$System_IO_MemoryStream = $null
			$form_Updates.Name = 'form_Updates'
			$form_Updates.Text = 'Application Update'
			$form_Updates.add_Load($form_Updates_Load)
			#
			# buttonDismiss
			#
			$buttonDismiss.Anchor = 'Bottom, Right'
			$buttonDismiss.Location = New-Object System.Drawing.Point(553, 283)
			$buttonDismiss.Margin = '5, 5, 5, 5'
			$buttonDismiss.Name = 'buttonDismiss'
			$buttonDismiss.Size = New-Object System.Drawing.Size(125, 35)
			$buttonDismiss.TabIndex = 4
			$buttonDismiss.Text = 'Dismiss'
			$buttonDismiss.UseVisualStyleBackColor = $True
			$buttonDismiss.add_Click($buttonDismiss_Click)
			#
			# linklabel_latestRelease
			#
			$linklabel_latestRelease.Anchor = 'Bottom, Left'
			$linklabel_latestRelease.AutoSize = $True
			$linklabel_latestRelease.Location = New-Object System.Drawing.Point(23, 285)
			$linklabel_latestRelease.Margin = '5, 0, 5, 0'
			$linklabel_latestRelease.Name = 'linklabel_latestRelease'
			$linklabel_latestRelease.Size = New-Object System.Drawing.Size(122, 20)
			$linklabel_latestRelease.TabIndex = 3
			$linklabel_latestRelease.TabStop = $True
			$linklabel_latestRelease.Text = 'Latest Release'
			$linklabel_latestRelease.add_LinkClicked($linklabel_latestRelease_LinkClicked)
			#
			# label_Title
			#
			$label_Title.AutoSize = $True
			$label_Title.Font = [System.Drawing.Font]::new('Microsoft Sans Serif', '11')
			$label_Title.Location = New-Object System.Drawing.Point(23, 19)
			$label_Title.Margin = '5, 0, 5, 0'
			$label_Title.Name = 'label_Title'
			$label_Title.Size = New-Object System.Drawing.Size(52, 26)
			$label_Title.TabIndex = 2
			$label_Title.Text = 'Title'
			#
			# label_ReleaseNotes
			#
			$label_ReleaseNotes.Anchor = 'Top, Bottom, Left'
			$label_ReleaseNotes.AutoSize = $True
			$label_ReleaseNotes.Location = New-Object System.Drawing.Point(23, 76)
			$label_ReleaseNotes.Margin = '5, 0, 5, 0'
			$label_ReleaseNotes.Name = 'label_ReleaseNotes'
			$label_ReleaseNotes.Size = New-Object System.Drawing.Size(124, 20)
			$label_ReleaseNotes.TabIndex = 1
			$label_ReleaseNotes.Text = 'Release Notes:'
			#
			# richtextbox1
			#
			$richtextbox1.Anchor = 'Top, Bottom, Left, Right'
			$richtextbox1.BackColor = [System.Drawing.SystemColors]::Control
			$richtextbox1.Cursor = 'Default'
			$richtextbox1.Location = New-Object System.Drawing.Point(23, 112)
			$richtextbox1.Margin = '5, 5, 5, 5'
			$richtextbox1.Name = 'richtextbox1'
			$richtextbox1.ReadOnly = $True
			$richtextbox1.Size = New-Object System.Drawing.Size(655, 161)
			$richtextbox1.TabIndex = 0
			$richtextbox1.Text = ''
			$form_Updates.ResumeLayout()
			#endregion Generated Form Code
			
			#----------------------------------------------
			
			#Save the initial state of the form
			$InitialFormWindowState = $form_Updates.WindowState
			#Init the OnLoad event to correct the initial state of the form
			$form_Updates.add_Load($Form_StateCorrection_Load)
			#Clean up the control events
			$form_Updates.add_FormClosed($Form_Cleanup_FormClosed)
			#Show the Form
			return $form_Updates.ShowDialog()
			
		}
		
		$Repo = $URI -replace ".*github.com/"
		
		try
		{
			# Attempt to get latest version from releases/latest url
			$LatestReleaseInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest" -Method Get -ErrorAction Stop
			
		}
		catch
		{
			$RepoReleaseInfo = Invoke-RestMethod -uri "https://api.github.com/repos/$Repo/releases" -Method Get
			$LatestReleaseInfo = $RepoReleaseInfo | Sort-Object -Property tag_name | Select-Object -Last 1
		}
		
		$ReleaseVersion = $($LatestReleaseInfo | Select-Object -ExpandProperty tag_name) -replace "v" -replace "-.*" # improve upon this later as its too simple atm
		
		if ($ReleaseVersion -gt $AppVersion)
		{
			Show-AppUpdaterPrompt -AppName $AppName -ReleaseNotes $LatestReleaseInfo.body -ReleaseURL $LatestReleaseInfo.html_url
		}
		else
		{
			if ($Startup -ne $true)
			{
				Add-Type -AssemblyName PresentationFramework
				$prompt = [System.Windows.MessageBox]::Show('No updates found', 'Update checker', 'OK', 'Information')
			}
		}
	}
	function Disable-ImportantControls
	{
		# This is more so a place holder for future stuff
		$button_Install.Enabled = $false
		$button_JustDownload.Enabled = $false
	}
	function Enable-ImportantControls
	{
	
		$button_Install.Enabled = $true
		$button_JustDownload.Enabled = $true
	}
	#region Variables
	$ApplicationVersion = "0.0.4"
	$OSInfo = Get-CimInstance CIM_OperatingSystem
	#endregion Variables
	
	$form_365toolkit_Load={
		Check-ApplicationUpdates -Startup -URI "https://github.com/serialscriptr/OfficeDeploymentGui" -AppVersion $ApplicationVersion -AppName "Office Toolkit GUI"
	}
	
	$uIFontSettingsToolStripMenuItem_Click={
		$fontdialog_main.ShowDialog()
	}
	
	$fontdialog_main_Apply={
		$FontVariables = Get-Variable | where { $_.Value.font }
		$ErrorActionPreference = 'SilentlyContinue'
		
		forEach ($Var in $FontVariables)
		{
			$NewVarValue = $Var.Value
			$NewVarValue.font = $fontdialog_main.Font
			$Var | Set-Variable -Value $NewVarValue
		}
		
		$ErrorActionPreference = 'Continue'
	}
	
	$combobox_Suite_SelectedIndexChanged={
		$Script:Suite = $combobox_Suite.Text
		
		switch ($suite) {
			"Office 365 Apps for Enterprise" {
				$IncludedApps = @("Access",
					"OneDrive (Groove)",
					"OneDrive Desktop",
					"Outlook",
					"Publisher",
					"Word",
					"Excel",
					"Skype for Business",
					"OneNote",
					"PowerPoint",
					"Teams")
				
				$DefaultChecked = @("Access",
					"OneDrive Desktop",
					"Outlook",
					"Publisher",
					"Word",
					"Excel",
					"OneNote",
					"PowerPoint")
	
				$listview_apps.Enabled = $true
				$listview_apps.Items.Clear()
				$listview_apps.Items.AddRange($IncludedApps)
				
				$listview_apps.Items | ForEach-Object {
					$item = $_
					
					if ($DefaultChecked -contains $item.text)
					{
						$item.Checked = $true
					}
					
					switch ($item.text) {
						"Teams" {
							$item.SubItems.Add('Use the MSI installer for better control')
						}
						"OneDrive (Groove)" {
							$item.SubItems.Add('Legacy version of OneDrive')
						}
					}
				}
			}
			"Office 365 Apps for Business" {
				$IncludedApps = @("Access",
					"OneDrive (Groove)",
					"OneDrive Desktop",
					"Outlook",
					"Publisher",
					"Word",
					"Excel",
					"Skype for Business",
					"OneNote",
					"PowerPoint",
					"Teams")
				
				$DefaultChecked = @("Access",
					"OneDrive Desktop",
					"Outlook",
					"Publisher",
					"Word",
					"Excel",
					"OneNote",
					"PowerPoint")
				
				
				$listview_apps.Enabled = $true
				$listview_apps.Items.Clear()
				$listview_apps.Items.AddRange($IncludedApps)
				
				$listview_apps.Items | ForEach-Object {
					$item = $_
					
					if ($DefaultChecked -contains $item.text)
					{
						$item.Checked = $true
					}
					
					switch ($item.text)
					{
						"Teams" {
							$item.SubItems.Add('Use the MSI installer for better control')
						}
						"OneDrive (Groove)" {
							$item.SubItems.Add('Legacy version of OneDrive')
						}
					}
				}
			}
			"Office LTSC Professional Plus 2021 - Volume License" {
				#<code>
			}
			"Office LTSC Professional Plus 2021 (SPLA) - Volume License" {
				#<code>
			}
			"Office LTSC Standard 2021 - Volume License"{
				
			}
			"Office LTSC Standard 2021 (SPLA) - Volume License"{
				
			}
			"Office Professional Plus 2019 - Volume License"{
				
			}
			"Office Standard 2019 - Volume License"{
				
			}
		}
	}
	
	$button_JustDownload_Click={
		$folderbrowsermoderndialog_main.ShowDialog()
		
		$DL_XML = Generate-DownloadOfficeXML -SourcePath $folderbrowsermoderndialog_main.SelectedPath
		$Toolkit_Exec = Download-ToolkitExec -Path $folderbrowsermoderndialog_main.SelectedPath
		
		New-Item -Path $folderbrowsermoderndialog_main.SelectedPath -ItemType File -Name "Office365_Download.xml" -Value $DL_XML -Force | Out-Null
		$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Office download xml created`n")
		$DL_XML_Path = "$($folderbrowsermoderndialog_main.SelectedPath)\Office365_Download.xml"
		$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Starting Office download...`n")
		
		$DLJobScript = {
			$process = Start-Process $Using:Toolkit_Exec -ArgumentList "/download `"$Using:DL_XML_Path`"" -Wait -PassThru
			Return $Process
		}
		$Script:JobNames = "DL-Job"
		Start-Job -Name "DL-Job" -ScriptBlock $DLJobScript
		Disable-ImportantControls
		$timer_JobTracker.Start()
		$Script:DLPath = $folderbrowsermoderndialog_main.SelectedPath
	}
	
	$button_Install_Click = {
		
		if ($DLPath)
		{
			$Install_XML = Generate-InstallOfficeXML
			
			New-Item -Path $DLPath -ItemType File -Name "Office365_Install.xml" -Value $Install_XML -Force | Out-Null
			$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Office install xml created`n")
			$Install_XML_Path = "$DLPath\Office365_Install.xml"
			$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Starting Office install...`n")
			
		}
		else
		{
			$folderbrowsermoderndialog_main.ShowDialog()
			
			$Install_XML = Generate-InstallOfficeXML
			$Toolkit_Exec = Download-ToolkitExec -Path $folderbrowsermoderndialog_main.SelectedPath
			
			New-Item -Path $($folderbrowsermoderndialog_main.SelectedPath) -ItemType File -Name "Office365_Install.xml" -Value $Install_XML -Force | Out-Null
			$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Office install xml created`n")
			$Install_XML_Path = "$($folderbrowsermoderndialog_main.SelectedPath)\Office365_Install.xml"
			$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Starting Office install...`n")
		}
		$InstallJobScipt = {
			$process = Start-Process $Using:Toolkit_Exec -ArgumentList "/configure `"$Using:Install_XML_Path`"" -Wait -PassThru
			Return $Process
		}
		$Script:JobNames = "In-Job"
		Start-Job -Name "In-Job" -ScriptBlock $InstallJobScipt
		Disable-ImportantControls
		$timer_JobTracker.Start()	
	}
		
	$linklabelLearnMoreAboutUserBa_LinkClicked=[System.Windows.Forms.LinkLabelLinkClickedEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.LinkLabelLinkClickedEventArgs]
		Start-Process "https://go.microsoft.com/fwlink/p/?linkid=2094112"
	}
	
	$linklabelLearnMoreAboutShareC_LinkClicked=[System.Windows.Forms.LinkLabelLinkClickedEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.LinkLabelLinkClickedEventArgs]
		Start-Process "https://docs.microsoft.com/en-us/deployoffice/overview-shared-computer-activation"
	}
	
	$linklabelLearnMoreAboutDevice_LinkClicked=[System.Windows.Forms.LinkLabelLinkClickedEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.LinkLabelLinkClickedEventArgs]
		Start-Process "https://docs.microsoft.com/en-us/deployoffice/overview-shared-computer-activation"
	}
	
	$exitToolStripMenuItem_Click={
		$form_365toolkit.Close()
	}
	
	$radiobuttonSharedComputer_CheckedChanged = {
		
		if ($this.checked)
		{
			$checkbox_SCLCacheOverride.Enabled = $true
		}
		else
		{
			$checkbox_SCLCacheOverride.Enabled = $false
			$textbox_TokenPath.Enabled = $false
		}
	}
	
	$checkbox_SCLCacheOverride_CheckedChanged={
		
		if ($this.checked)
		{
			$textbox_TokenPath.Enabled = $true
		}
		else
		{
			$textbox_TokenPath.Enabled = $false
		}
	}
	
	$checkForUpdatesToolStripMenuItem_Click={
		Check-ApplicationUpdates -URI "https://github.com/serialscriptr/OfficeDeploymentGui" -AppVersion $ApplicationVersion -AppName "Office Toolkit GUI"
	}
	
	$installXMLToolStripMenuItem_Click={
		$savefiledialog_main.ShowDialog()
		
		$Xml = Generate-InstallOfficeXML
		
		New-Item -Path $savefiledialog_main.FileName -Value $Xml -Force
	}
	
	$downloadXMLToolStripMenuItem_Click={
		$savefiledialog_main.ShowDialog()
		
		$xml = Generate-DownloadOfficeXML
		
		New-Item -Path $savefiledialog_main.FileName -Value $xml -Force
	}
	
	$timer_JobTracker_Tick = {
		$timer_JobTracker.Stop()
		$Jobs = Get-job | Where-Object { $JobNames -icontains $_.name }
		
		forEach ($Job in $Jobs)
		{
			switch ($Job.Name) {
				"DL-Job" {
					if ($Job.State -eq "Completed")
					{
						$Process = Receive-Job -Job $Job
						$TimeElapsed = New-TimeSpan -Start $process.StartTime -End $process.ExitTime
						$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Download complete, time elapsed $("$("{0:D2}" -f $TimeElapsed.Hours):$("{0:D2}" -f $TimeElapsed.Minutes):$("{0:D2}" -f $TimeElapsed.Seconds)")`n")
						Remove-Job -Job $Job
					}
				}
				"In-Job"{
					if ($Job.State -eq "Completed")
					{
						$Process = Receive-Job -Job $Job
						$TimeElapsed = New-TimeSpan -Start $process.StartTime -End $process.ExitTime
						$richtextbox_ProcessInfo.AppendText("$(Get-Date -Format "HH:mm:ss") : Install complete, time elapsed $("$("{0:D2}" -f $TimeElapsed.Hours):$("{0:D2}" -f $TimeElapsed.Minutes):$("{0:D2}" -f $TimeElapsed.Seconds)")`n")
						Remove-Job -Job $Job
					}
				}
			}
		}
		
		$Jobs = Get-job | Where-Object { $JobNames -icontains $_.name }
		$boolCheck = [bool]$Jobs
		
		if ($boolCheck)
		{
			$timer_JobTracker.Start()
		}
		else
		{
			Enable-ImportantControls	
		}
	}
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$form_365toolkit.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$combobox_suite.remove_SelectedIndexChanged($combobox_Suite_SelectedIndexChanged)
			$linklabelLearnMoreAboutDevice.remove_LinkClicked($linklabelLearnMoreAboutDevice_LinkClicked)
			$checkbox_SCLCacheOverride.remove_CheckedChanged($checkbox_SCLCacheOverride_CheckedChanged)
			$linklabelLearnMoreAboutShareC.remove_LinkClicked($linklabelLearnMoreAboutShareC_LinkClicked)
			$linklabelLearnMoreAboutUserBa.remove_LinkClicked($linklabelLearnMoreAboutUserBa_LinkClicked)
			$radiobuttonSharedComputer.remove_CheckedChanged($radiobuttonSharedComputer_CheckedChanged)
			$form_365toolkit.remove_Load($form_365toolkit_Load)
			$button_Install.remove_Click($button_Install_Click)
			$button_JustDownload.remove_Click($button_JustDownload_Click)
			$uIFontSettingsToolStripMenuItem.remove_Click($uIFontSettingsToolStripMenuItem_Click)
			$fontdialog_main.remove_Apply($fontdialog_main_Apply)
			$exitToolStripMenuItem.remove_Click($exitToolStripMenuItem_Click)
			$checkForUpdatesToolStripMenuItem.remove_Click($checkForUpdatesToolStripMenuItem_Click)
			$installXMLToolStripMenuItem.remove_Click($installXMLToolStripMenuItem_Click)
			$downloadXMLToolStripMenuItem.remove_Click($downloadXMLToolStripMenuItem_Click)
			$timer_JobTracker.remove_Tick($timer_JobTracker_Tick)
			$form_365toolkit.remove_Load($Form_StateCorrection_Load)
			$form_365toolkit.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$form_365toolkit.SuspendLayout()
	$tabcontrol1.SuspendLayout()
	$tabpage_Suite.SuspendLayout()
	$groupbox1.SuspendLayout()
	$tabpage2.SuspendLayout()
	$groupbox_Activation.SuspendLayout()
	$tabpage1.SuspendLayout()
	$splitcontainer1.BeginInit()
	$splitcontainer1.SuspendLayout()
	$menustrip_main.SuspendLayout()
	#
	# form_365toolkit
	#
	$form_365toolkit.Controls.Add($tabcontrol1)
	$form_365toolkit.Controls.Add($menustrip_main)
	$form_365toolkit.AutoScaleDimensions = New-Object System.Drawing.SizeF(10, 20)
	$form_365toolkit.AutoScaleMode = 'Font'
	$form_365toolkit.AutoScroll = $True
	$form_365toolkit.AutoSizeMode = 'GrowAndShrink'
	$form_365toolkit.BackColor = [System.Drawing.SystemColors]::ControlLightLight 
	$form_365toolkit.ClientSize = New-Object System.Drawing.Size(774, 532)
	#region Binary Data
	$Formatter_binaryFomatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
	$System_IO_MemoryStream = New-Object System.IO.MemoryStream (,[byte[]][System.Convert]::FromBase64String('
AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj00LjAuMC4wLCBD
dWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABNTeXN0
ZW0uRHJhd2luZy5JY29uAgAAAAhJY29uRGF0YQhJY29uU2l6ZQcEAhNTeXN0ZW0uRHJhd2luZy5T
aXplAgAAAAIAAAAJAwAAAAX8////E1N5c3RlbS5EcmF3aW5nLlNpemUCAAAABXdpZHRoBmhlaWdo
dAAACAgCAAAAAAAAAAAAAAAPAwAAACM9AAACAAABAAEAAAAAAAEAIAANPQAAFgAAAIlQTkcNChoK
AAAADUlIRFIAAAEAAAABAAgGAAAAXHKoZgAAAAFvck5UAc+id5oAADzHSURBVHja7Z0HeBTVGobP
pPfey2yS3VR6h4QOgoioiCBi7w0VC/ZekQ4ivUhVUKQo0quIiu2KFQTNBARFFAFBFPXc78zOJpPN
brK72WQ37P8/z/sM6nPv9c7O955/zj+7wxgVFRUVFRUVFRUVFRUVFRUVFRUVFdXZU3f6+bHzI/zZ
TYH+7PogP5VrwFVBErsSXBXox64Wfw6U2LX4s4CKiqqhBj7En90c7M9uCTYfbwwyB39AKBPHIBAD
ASRBACkQQDIEkHB1oBQO/AdDAkPANeESe6oxY3wlY3P60DmlovL6GhocgND7sdvUoz/r7icJASRC
AB0R+tvB5GuD/FYh/DuvDvL7CgLYAwF8C3ZdEShtvTxQeg08DQFcAgEUjGjFwpowiU3tKbFDL0EG
Kxj7awFjZzbSuaai8poaFhrAbgoKYHeEBLDbEX4IIPLWYP9eCP/Em4L9/3djsP/xG4L8OQTArwOQ
AIcEOCTAr1KROCTAIQEOAXAI4MzlQdLBayOlDffI0uPzLpC6nl7Akjhn0vFZjP290CyCvxaBhXT+
qag8UsNDgthdIYFsGBDhB7HoAq6CADZCAH+gC+CQAIcEOCTAbwiukMB1OglcbVsC/DLx52CJ35wg
/TGmWNr10f1sJsJ/FX+dFUEAYX/OY+z0/Ar+fBVCmEufCxVVndb9/n7sXoT+PnA3uCs0MOSukID+
EMDmoSEBf6ML4JAAvxXoJVAugmq6gSttiGBwAMBxaKrEXx3A/j00kR08s4itgwSeQPB7gBR0CH6n
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
QmCCCw=='))
	#endregion
	$form_365toolkit.Icon = $Formatter_binaryFomatter.Deserialize($System_IO_MemoryStream)
	$Formatter_binaryFomatter = $null
	$System_IO_MemoryStream = $null
	$form_365toolkit.MainMenuStrip = $menustrip_main
	$form_365toolkit.Margin = '4, 4, 4, 4'
	$form_365toolkit.Name = 'form_365toolkit'
	$form_365toolkit.Text = '365 Toolkit GUI'
	$form_365toolkit.add_Load($form_365toolkit_Load)
	#
	# tabcontrol1
	#
	$tabcontrol1.Controls.Add($tabpage_Suite)
	$tabcontrol1.Controls.Add($tabpage2)
	$tabcontrol1.Controls.Add($tabpage1)
	$tabcontrol1.Dock = 'Fill'
	$tabcontrol1.Location = New-Object System.Drawing.Point(0, 37)
	$tabcontrol1.Margin = '5, 5, 5, 5'
	$tabcontrol1.Name = 'tabcontrol1'
	$tabcontrol1.SelectedIndex = 0
	$tabcontrol1.Size = New-Object System.Drawing.Size(774, 495)
	$tabcontrol1.TabIndex = 1
	#
	# tabpage_Suite
	#
	$tabpage_Suite.Controls.Add($listview_apps)
	$tabpage_Suite.Controls.Add($combobox_suite)
	$tabpage_Suite.Controls.Add($labelIncludedApps)
	$tabpage_Suite.Controls.Add($labelOfficeSuite)
	$tabpage_Suite.Controls.Add($groupbox1)
	$tabpage_Suite.Location = New-Object System.Drawing.Point(4, 29)
	$tabpage_Suite.Margin = '5, 5, 5, 5'
	$tabpage_Suite.Name = 'tabpage_Suite'
	$tabpage_Suite.Padding = '2, 4, 2, 4'
	$tabpage_Suite.Size = New-Object System.Drawing.Size(766, 462)
	$tabpage_Suite.TabIndex = 0
	$tabpage_Suite.Text = 'Suite'
	$tabpage_Suite.UseVisualStyleBackColor = $True
	#
	# listview_apps
	#
	$listview_apps.Anchor = 'Top, Bottom, Left, Right'
	$listview_apps.CheckBoxes = $True
	[void]$listview_apps.Columns.Add($columnheader1)
	[void]$listview_apps.Columns.Add($columnheader2)
	$listview_apps.Enabled = $False
	$listview_apps.GridLines = $True
	$listview_apps.HideSelection = $False
	$listview_apps.Location = New-Object System.Drawing.Point(198, 149)
	$listview_apps.Margin = '5, 5, 5, 5'
	$listview_apps.Name = 'listview_apps'
	$listview_apps.Size = New-Object System.Drawing.Size(558, 299)
	$listview_apps.TabIndex = 15
	$listview_apps.UseCompatibleStateImageBehavior = $False
	$listview_apps.View = 'Details'
	#
	# combobox_suite
	#
	$combobox_suite.Anchor = 'Top, Left, Right'
	$combobox_suite.FormattingEnabled = $True
	[void]$combobox_suite.Items.Add('Office 365 Apps for Enterprise')
	[void]$combobox_suite.Items.Add('Office 365 Apps for Business')
	$combobox_suite.Location = New-Object System.Drawing.Point(198, 56)
	$combobox_suite.Margin = '5, 5, 5, 5'
	$combobox_suite.Name = 'combobox_suite'
	$combobox_suite.Size = New-Object System.Drawing.Size(559, 28)
	$combobox_suite.TabIndex = 14
	$combobox_suite.add_SelectedIndexChanged($combobox_Suite_SelectedIndexChanged)
	#
	# labelIncludedApps
	#
	$labelIncludedApps.Anchor = 'Top, Bottom, Left, Right'
	$labelIncludedApps.AutoSize = $True
	$labelIncludedApps.Location = New-Object System.Drawing.Point(58, 159)
	$labelIncludedApps.Margin = '5, 0, 5, 0'
	$labelIncludedApps.Name = 'labelIncludedApps'
	$labelIncludedApps.Size = New-Object System.Drawing.Size(114, 20)
	$labelIncludedApps.TabIndex = 13
	$labelIncludedApps.Text = 'Included Apps'
	#
	# labelOfficeSuite
	#
	$labelOfficeSuite.Anchor = 'Top, Bottom, Left, Right'
	$labelOfficeSuite.AutoSize = $True
	$labelOfficeSuite.Location = New-Object System.Drawing.Point(192, 21)
	$labelOfficeSuite.Margin = '5, 0, 5, 0'
	$labelOfficeSuite.Name = 'labelOfficeSuite'
	$labelOfficeSuite.Size = New-Object System.Drawing.Size(102, 20)
	$labelOfficeSuite.TabIndex = 5
	$labelOfficeSuite.Text = 'Office Suite:'
	#
	# groupbox1
	#
	$groupbox1.Controls.Add($radiobutton_64bit)
	$groupbox1.Controls.Add($radiobutton_32bit)
	$groupbox1.Location = New-Object System.Drawing.Point(8, 8)
	$groupbox1.Margin = '5, 5, 5, 5'
	$groupbox1.Name = 'groupbox1'
	$groupbox1.Padding = '5, 5, 5, 5'
	$groupbox1.Size = New-Object System.Drawing.Size(174, 118)
	$groupbox1.TabIndex = 3
	$groupbox1.TabStop = $False
	$groupbox1.Text = 'Architecture'
	$tooltip_main.SetToolTip($groupbox1, 'Architecture of deployed office apps')
	#
	# radiobutton_64bit
	#
	$radiobutton_64bit.Location = New-Object System.Drawing.Point(16, 71)
	$radiobutton_64bit.Margin = '5, 5, 5, 5'
	$radiobutton_64bit.Name = 'radiobutton_64bit'
	$radiobutton_64bit.Size = New-Object System.Drawing.Size(148, 36)
	$radiobutton_64bit.TabIndex = 1
	$radiobutton_64bit.Text = '64-bit'
	$tooltip_main.SetToolTip($radiobutton_64bit, 'Select 64 bit office for best performance')
	$radiobutton_64bit.UseVisualStyleBackColor = $True
	#
	# radiobutton_32bit
	#
	$radiobutton_32bit.Checked = $True
	$radiobutton_32bit.Location = New-Object System.Drawing.Point(16, 29)
	$radiobutton_32bit.Margin = '5, 5, 5, 5'
	$radiobutton_32bit.Name = 'radiobutton_32bit'
	$radiobutton_32bit.Size = New-Object System.Drawing.Size(148, 36)
	$radiobutton_32bit.TabIndex = 0
	$radiobutton_32bit.TabStop = $True
	$radiobutton_32bit.Text = '32-bit'
	$tooltip_main.SetToolTip($radiobutton_32bit, 'Select 32 bit for better compatibility with office addons')
	$radiobutton_32bit.UseVisualStyleBackColor = $True
	#
	# tabpage2
	#
	$tabpage2.Controls.Add($groupbox_Activation)
	$tabpage2.Controls.Add($richtextbox_eula)
	$tabpage2.Controls.Add($checkbox_AcceptEULA)
	$tabpage2.Location = New-Object System.Drawing.Point(4, 29)
	$tabpage2.Margin = '5, 5, 5, 5'
	$tabpage2.Name = 'tabpage2'
	$tabpage2.Padding = '2, 4, 2, 4'
	$tabpage2.Size = New-Object System.Drawing.Size(766, 462)
	$tabpage2.TabIndex = 1
	$tabpage2.Text = 'Licensing'
	$tabpage2.UseVisualStyleBackColor = $True
	#
	# groupbox_Activation
	#
	$groupbox_Activation.Controls.Add($linklabelLearnMoreAboutDevice)
	$groupbox_Activation.Controls.Add($radiobutton_deviceBased)
	$groupbox_Activation.Controls.Add($textbox_TokenPath)
	$groupbox_Activation.Controls.Add($checkbox_SCLCacheOverride)
	$groupbox_Activation.Controls.Add($linklabelLearnMoreAboutShareC)
	$groupbox_Activation.Controls.Add($linklabelLearnMoreAboutUserBa)
	$groupbox_Activation.Controls.Add($radiobuttonSharedComputer)
	$groupbox_Activation.Controls.Add($radiobuttonUserBased)
	$groupbox_Activation.Anchor = 'Top, Bottom, Left, Right'
	$groupbox_Activation.Location = New-Object System.Drawing.Point(10, 221)
	$groupbox_Activation.Margin = '5, 5, 5, 5'
	$groupbox_Activation.Name = 'groupbox_Activation'
	$groupbox_Activation.Padding = '5, 5, 5, 5'
	$groupbox_Activation.Size = New-Object System.Drawing.Size(751, 231)
	$groupbox_Activation.TabIndex = 5
	$groupbox_Activation.TabStop = $False
	$groupbox_Activation.Text = 'Product Activation'
	#
	# linklabelLearnMoreAboutDevice
	#
	$linklabelLearnMoreAboutDevice.Anchor = 'Top, Right'
	$linklabelLearnMoreAboutDevice.AutoSize = $True
	$linklabelLearnMoreAboutDevice.Location = New-Object System.Drawing.Point(398, 195)
	$linklabelLearnMoreAboutDevice.Margin = '5, 0, 5, 0'
	$linklabelLearnMoreAboutDevice.Name = 'linklabelLearnMoreAboutDevice'
	$linklabelLearnMoreAboutDevice.Size = New-Object System.Drawing.Size(315, 20)
	$linklabelLearnMoreAboutDevice.TabIndex = 7
	$linklabelLearnMoreAboutDevice.TabStop = $True
	$linklabelLearnMoreAboutDevice.Text = 'Learn more about device based licensing'
	$linklabelLearnMoreAboutDevice.add_LinkClicked($linklabelLearnMoreAboutDevice_LinkClicked)
	#
	# radiobutton_deviceBased
	#
	$radiobutton_deviceBased.AutoSize = $True
	$radiobutton_deviceBased.Enabled = $False
	$radiobutton_deviceBased.Location = New-Object System.Drawing.Point(10, 193)
	$radiobutton_deviceBased.Margin = '5, 5, 5, 5'
	$radiobutton_deviceBased.Name = 'radiobutton_deviceBased'
	$radiobutton_deviceBased.Size = New-Object System.Drawing.Size(139, 24)
	$radiobutton_deviceBased.TabIndex = 6
	$radiobutton_deviceBased.Text = 'Device Based'
	$radiobutton_deviceBased.UseVisualStyleBackColor = $True
	#
	# textbox_TokenPath
	#
	$textbox_TokenPath.Anchor = 'Top, Left, Right'
	$textbox_TokenPath.Enabled = $False
	$textbox_TokenPath.Location = New-Object System.Drawing.Point(35, 156)
	$textbox_TokenPath.Margin = '5, 5, 5, 5'
	$textbox_TokenPath.Name = 'textbox_TokenPath'
	$textbox_TokenPath.Size = New-Object System.Drawing.Size(678, 26)
	$textbox_TokenPath.TabIndex = 5
	$tooltip_main.SetToolTip($textbox_TokenPath, 'Network, local, or HTTP path')
	#
	# checkbox_SCLCacheOverride
	#
	$checkbox_SCLCacheOverride.Enabled = $False
	$checkbox_SCLCacheOverride.Location = New-Object System.Drawing.Point(35, 120)
	$checkbox_SCLCacheOverride.Margin = '5, 5, 5, 5'
	$checkbox_SCLCacheOverride.Name = 'checkbox_SCLCacheOverride'
	$checkbox_SCLCacheOverride.Size = New-Object System.Drawing.Size(362, 27)
	$checkbox_SCLCacheOverride.TabIndex = 4
	$checkbox_SCLCacheOverride.Text = 'Allow the licensing token to roam'
	$checkbox_SCLCacheOverride.UseVisualStyleBackColor = $True
	$checkbox_SCLCacheOverride.add_CheckedChanged($checkbox_SCLCacheOverride_CheckedChanged)
	#
	# linklabelLearnMoreAboutShareC
	#
	$linklabelLearnMoreAboutShareC.Anchor = 'Top, Right'
	$linklabelLearnMoreAboutShareC.AutoSize = $True
	$linklabelLearnMoreAboutShareC.Location = New-Object System.Drawing.Point(379, 88)
	$linklabelLearnMoreAboutShareC.Margin = '5, 0, 5, 0'
	$linklabelLearnMoreAboutShareC.Name = 'linklabelLearnMoreAboutShareC'
	$linklabelLearnMoreAboutShareC.Size = New-Object System.Drawing.Size(339, 20)
	$linklabelLearnMoreAboutShareC.TabIndex = 3
	$linklabelLearnMoreAboutShareC.TabStop = $True
	$linklabelLearnMoreAboutShareC.Text = 'Learn more about share computer activation'
	$linklabelLearnMoreAboutShareC.add_LinkClicked($linklabelLearnMoreAboutShareC_LinkClicked)
	#
	# linklabelLearnMoreAboutUserBa
	#
	$linklabelLearnMoreAboutUserBa.Anchor = 'Top, Right'
	$linklabelLearnMoreAboutUserBa.AutoSize = $True
	$linklabelLearnMoreAboutUserBa.Location = New-Object System.Drawing.Point(418, 41)
	$linklabelLearnMoreAboutUserBa.Margin = '5, 0, 5, 0'
	$linklabelLearnMoreAboutUserBa.Name = 'linklabelLearnMoreAboutUserBa'
	$linklabelLearnMoreAboutUserBa.Size = New-Object System.Drawing.Size(300, 20)
	$linklabelLearnMoreAboutUserBa.TabIndex = 2
	$linklabelLearnMoreAboutUserBa.TabStop = $True
	$linklabelLearnMoreAboutUserBa.Text = 'Learn more about user based licensing'
	$linklabelLearnMoreAboutUserBa.add_LinkClicked($linklabelLearnMoreAboutUserBa_LinkClicked)
	#
	# radiobuttonSharedComputer
	#
	$radiobuttonSharedComputer.AutoSize = $True
	$radiobuttonSharedComputer.Location = New-Object System.Drawing.Point(10, 86)
	$radiobuttonSharedComputer.Margin = '5, 5, 5, 5'
	$radiobuttonSharedComputer.Name = 'radiobuttonSharedComputer'
	$radiobuttonSharedComputer.Size = New-Object System.Drawing.Size(165, 24)
	$radiobuttonSharedComputer.TabIndex = 1
	$radiobuttonSharedComputer.Text = 'Shared Computer'
	$radiobuttonSharedComputer.UseVisualStyleBackColor = $True
	$radiobuttonSharedComputer.add_CheckedChanged($radiobuttonSharedComputer_CheckedChanged)
	#
	# radiobuttonUserBased
	#
	$radiobuttonUserBased.AutoSize = $True
	$radiobuttonUserBased.Checked = $True
	$radiobuttonUserBased.Location = New-Object System.Drawing.Point(10, 39)
	$radiobuttonUserBased.Margin = '5, 5, 5, 5'
	$radiobuttonUserBased.Name = 'radiobuttonUserBased'
	$radiobuttonUserBased.Size = New-Object System.Drawing.Size(120, 24)
	$radiobuttonUserBased.TabIndex = 0
	$radiobuttonUserBased.TabStop = $True
	$radiobuttonUserBased.Text = 'User based'
	$radiobuttonUserBased.UseVisualStyleBackColor = $True
	#
	# richtextbox_eula
	#
	$richtextbox_eula.Anchor = 'Top, Left, Right'
	$richtextbox_eula.Cursor = 'Arrow'
	$richtextbox_eula.Location = New-Object System.Drawing.Point(40, 55)
	$richtextbox_eula.Margin = '5, 5, 5, 5'
	$richtextbox_eula.Name = 'richtextbox_eula'
	$richtextbox_eula.Size = New-Object System.Drawing.Size(688, 145)
	$richtextbox_eula.TabIndex = 4
	$richtextbox_eula.Text = 'Your use of the subscription service and software is subject to the terms and conditions of the agreement you agreed to when you signed up for the subscription and by which you acquired a license for the software. For instance, if you are:

• a volume license customer, use of this software is subject to your volume license agreement.
• a Microsoft Online Subscription customer, use of this software is subject to the Microsoft Online Subscription agreement.

You may not use the service or software if you have not validly acquired a license from Microsoft or its licensed distributors.'
	#
	# checkbox_AcceptEULA
	#
	$checkbox_AcceptEULA.AutoSize = $True
	$checkbox_AcceptEULA.Checked = $True
	$checkbox_AcceptEULA.CheckState = 'Checked'
	$checkbox_AcceptEULA.Location = New-Object System.Drawing.Point(10, 21)
	$checkbox_AcceptEULA.Margin = '5, 5, 5, 5'
	$checkbox_AcceptEULA.Name = 'checkbox_AcceptEULA'
	$checkbox_AcceptEULA.Size = New-Object System.Drawing.Size(267, 24)
	$checkbox_AcceptEULA.TabIndex = 3
	$checkbox_AcceptEULA.Text = 'Automatically accept the EULA'
	$checkbox_AcceptEULA.UseVisualStyleBackColor = $True
	#
	# tabpage1
	#
	$tabpage1.Controls.Add($splitcontainer1)
	$tabpage1.Location = New-Object System.Drawing.Point(4, 29)
	$tabpage1.Margin = '5, 5, 5, 5'
	$tabpage1.Name = 'tabpage1'
	$tabpage1.Padding = '2, 4, 2, 4'
	$tabpage1.Size = New-Object System.Drawing.Size(766, 462)
	$tabpage1.TabIndex = 2
	$tabpage1.Text = 'Process'
	$tabpage1.UseVisualStyleBackColor = $True
	#
	# splitcontainer1
	#
	$splitcontainer1.Dock = 'Fill'
	$splitcontainer1.Location = New-Object System.Drawing.Point(2, 4)
	$splitcontainer1.Margin = '5, 5, 5, 5'
	$splitcontainer1.Name = 'splitcontainer1'
	$splitcontainer1.Orientation = 'Horizontal'
	[void]$splitcontainer1.Panel1.Controls.Add($richtextbox_ProcessInfo)
	[void]$splitcontainer1.Panel2.Controls.Add($button_JustDownload)
	[void]$splitcontainer1.Panel2.Controls.Add($button_Install)
	$splitcontainer1.Size = New-Object System.Drawing.Size(762, 454)
	$splitcontainer1.SplitterDistance = 354
	$splitcontainer1.SplitterWidth = 7
	$splitcontainer1.TabIndex = 0
	#
	# menustrip_main
	#
	$menustrip_main.ImageScalingSize = New-Object System.Drawing.Size(24, 24)
	[void]$menustrip_main.Items.Add($fileToolStripMenuItem)
	[void]$menustrip_main.Items.Add($viewToolStripMenuItem)
	[void]$menustrip_main.Items.Add($aboutToolStripMenuItem)
	$menustrip_main.Location = New-Object System.Drawing.Point(0, 0)
	$menustrip_main.Name = 'menustrip_main'
	$menustrip_main.Padding = '10, 4, 0, 4'
	$menustrip_main.Size = New-Object System.Drawing.Size(774, 37)
	$menustrip_main.TabIndex = 0
	$menustrip_main.Text = 'menustrip1'
	#
	# fileToolStripMenuItem
	#
	[void]$fileToolStripMenuItem.DropDownItems.Add($exportCurrentXMLToolStripMenuItem)
	[void]$fileToolStripMenuItem.DropDownItems.Add($exitToolStripMenuItem)
	$fileToolStripMenuItem.Name = 'fileToolStripMenuItem'
	$fileToolStripMenuItem.Size = New-Object System.Drawing.Size(50, 29)
	$fileToolStripMenuItem.Text = 'File'
	#
	# viewToolStripMenuItem
	#
	[void]$viewToolStripMenuItem.DropDownItems.Add($uIFontSettingsToolStripMenuItem)
	$viewToolStripMenuItem.Name = 'viewToolStripMenuItem'
	$viewToolStripMenuItem.Size = New-Object System.Drawing.Size(61, 29)
	$viewToolStripMenuItem.Text = 'View'
	#
	# button_Install
	#
	$button_Install.Anchor = 'Top, Bottom, Right'
	$button_Install.Location = New-Object System.Drawing.Point(604, 18)
	$button_Install.Margin = '5, 5, 5, 5'
	$button_Install.Name = 'button_Install'
	$button_Install.Size = New-Object System.Drawing.Size(142, 43)
	$button_Install.TabIndex = 0
	$button_Install.Text = 'Install'
	$tooltip_main.SetToolTip($button_Install, 'Install the current configuration of Office to this device.
If it hasn''t already, Office will be downloaded.')
	$button_Install.UseVisualStyleBackColor = $True
	$button_Install.add_Click($button_Install_Click)
	#
	# button_JustDownload
	#
	$button_JustDownload.Anchor = 'Top, Bottom, Right'
	$button_JustDownload.Location = New-Object System.Drawing.Point(452, 18)
	$button_JustDownload.Margin = '5, 5, 5, 5'
	$button_JustDownload.Name = 'button_JustDownload'
	$button_JustDownload.Size = New-Object System.Drawing.Size(142, 43)
	$button_JustDownload.TabIndex = 1
	$button_JustDownload.Text = 'Just Download'
	$tooltip_main.SetToolTip($button_JustDownload, 'Downloads the selected Office configuration install files to the selected destication folder')
	$button_JustDownload.UseVisualStyleBackColor = $True
	$button_JustDownload.add_Click($button_JustDownload_Click)
	#
	# uIFontSettingsToolStripMenuItem
	#
	$uIFontSettingsToolStripMenuItem.Name = 'uIFontSettingsToolStripMenuItem'
	$uIFontSettingsToolStripMenuItem.Size = New-Object System.Drawing.Size(223, 30)
	$uIFontSettingsToolStripMenuItem.Text = 'UI Font Settings'
	$uIFontSettingsToolStripMenuItem.add_Click($uIFontSettingsToolStripMenuItem_Click)
	#
	# fontdialog_main
	#
	$fontdialog_main.ShowApply = $True
	$fontdialog_main.add_Apply($fontdialog_main_Apply)
	#
	# tooltip_main
	#
	$tooltip_main.ToolTipIcon = 'Info'
	$tooltip_main.ToolTipTitle = 'Info'
	#
	# folderbrowsermoderndialog_main
	#
	$folderbrowsermoderndialog_main.Title = 'Select a destination folder'
	#
	# richtextbox_ProcessInfo
	#
	$richtextbox_ProcessInfo.Anchor = 'Top, Bottom, Left, Right'
	$richtextbox_ProcessInfo.Location = New-Object System.Drawing.Point(19, 19)
	$richtextbox_ProcessInfo.Margin = '5, 5, 5, 5'
	$richtextbox_ProcessInfo.Name = 'richtextbox_ProcessInfo'
	$richtextbox_ProcessInfo.Size = New-Object System.Drawing.Size(727, 311)
	$richtextbox_ProcessInfo.TabIndex = 1
	$richtextbox_ProcessInfo.Text = ''
	#
	# exitToolStripMenuItem
	#
	$exitToolStripMenuItem.Name = 'exitToolStripMenuItem'
	$exitToolStripMenuItem.Size = New-Object System.Drawing.Size(250, 30)
	$exitToolStripMenuItem.Text = 'Exit'
	$exitToolStripMenuItem.add_Click($exitToolStripMenuItem_Click)
	#
	# aboutToolStripMenuItem
	#
	[void]$aboutToolStripMenuItem.DropDownItems.Add($checkForUpdatesToolStripMenuItem)
	$aboutToolStripMenuItem.Name = 'aboutToolStripMenuItem'
	$aboutToolStripMenuItem.Size = New-Object System.Drawing.Size(88, 29)
	$aboutToolStripMenuItem.Text = 'Options'
	#
	# checkForUpdatesToolStripMenuItem
	#
	$checkForUpdatesToolStripMenuItem.Name = 'checkForUpdatesToolStripMenuItem'
	$checkForUpdatesToolStripMenuItem.Size = New-Object System.Drawing.Size(242, 30)
	$checkForUpdatesToolStripMenuItem.Text = 'Check for Updates'
	$checkForUpdatesToolStripMenuItem.add_Click($checkForUpdatesToolStripMenuItem_Click)
	#
	# exportCurrentXMLToolStripMenuItem
	#
	[void]$exportCurrentXMLToolStripMenuItem.DropDownItems.Add($installXMLToolStripMenuItem)
	[void]$exportCurrentXMLToolStripMenuItem.DropDownItems.Add($downloadXMLToolStripMenuItem)
	$exportCurrentXMLToolStripMenuItem.Name = 'exportCurrentXMLToolStripMenuItem'
	$exportCurrentXMLToolStripMenuItem.Size = New-Object System.Drawing.Size(250, 30)
	$exportCurrentXMLToolStripMenuItem.Text = 'Export Current XML'
	#
	# installXMLToolStripMenuItem
	#
	$installXMLToolStripMenuItem.Name = 'installXMLToolStripMenuItem'
	$installXMLToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 30)
	$installXMLToolStripMenuItem.Text = 'Install XML'
	$installXMLToolStripMenuItem.add_Click($installXMLToolStripMenuItem_Click)
	#
	# downloadXMLToolStripMenuItem
	#
	$downloadXMLToolStripMenuItem.Name = 'downloadXMLToolStripMenuItem'
	$downloadXMLToolStripMenuItem.Size = New-Object System.Drawing.Size(218, 30)
	$downloadXMLToolStripMenuItem.Text = 'Download XML'
	$downloadXMLToolStripMenuItem.add_Click($downloadXMLToolStripMenuItem_Click)
	#
	# savefiledialog_main
	#
	$savefiledialog_main.DefaultExt = 'xml'
	#
	# columnheader1
	#
	$columnheader1.Text = 'Name'
	$columnheader1.Width = 240
	#
	# columnheader2
	#
	$columnheader2.Text = 'Info'
	$columnheader2.Width = 240
	#
	# timer_JobTracker
	#
	$timer_JobTracker.add_Tick($timer_JobTracker_Tick)
	$menustrip_main.ResumeLayout()
	$splitcontainer1.EndInit()
	$splitcontainer1.ResumeLayout()
	$tabpage1.ResumeLayout()
	$groupbox_Activation.ResumeLayout()
	$tabpage2.ResumeLayout()
	$groupbox1.ResumeLayout()
	$tabpage_Suite.ResumeLayout()
	$tabcontrol1.ResumeLayout()
	$form_365toolkit.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $form_365toolkit.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form_365toolkit.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$form_365toolkit.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $form_365toolkit.ShowDialog()

} #End Function

#Call the form
Show-365-Toolkit_psf | Out-Null