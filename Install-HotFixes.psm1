Function Install-HotFixes {
<#
.SYNOPSIS  
		Install Office Hotfixes with exit codes

.DESCRIPTION  
		Install Office Hotfixes with exit codes

.LINK  
    https://github.com/tomarbuthnot/Install-HotFixes
                
.NOTES  
	Version:
			0.2
  
	Author/Copyright:	 
			Copyright Tom Arbuthnot - All Rights Reserved
    
	Email/Blog/Twitter:	
			tom@tomarbuthnot.com tomtalks.uk @tomarbuthnot
    
	Disclaimer:   	
			THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
			OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
			While these scripts are tested and working in my environment, it is recommended 
			that you test these scripts in a test environment before using in your production 
			environment. Tom Arbuthnot further disclaims all implied warranties including, 
			without limitation, any implied warranties of merchantability or of fitness for 
			a particular purpose. The entire risk arising out of the use or performance of 
			this script and documentation remains with you. In no event shall Tom Arbuthnot, 
			its authors, or anyone else involved in the creation, production, or delivery of 
			this script/tool be liable for any damages whatsoever (including, without limitation, 
			damages for loss of business profits, business interruption, loss of business 
			information, or other pecuniary loss) arising out of the use of or inability to use 
			the sample scripts or documentation, even if Tom Arbuthnot has been advised of 
			the possibility of such damages.
    
     
	Acknowledgements: 	
    
	Assumptions:	      
			ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    
	Limitations:		  										
    		
	Ideas/Wish list:
			Detects loops based on number of log lines, this can vary, should detect based on timestamp of line 
    
	Rights Required:	

	Known issues:	


.EXAMPLE
		Function-Template
 
		Description
		-----------
		Returns Objects

.EXAMPLE
		Function-Template -Param1
 
		Description
		-----------
		Actions Param1

# Parameters

.PARAMETER Param1
		Param1 description

.PARAMETER Param2
		Param2 Description
		

#>
  
  
  #############################################################
  # Param Block
  #############################################################
  
  # Sets that -Whatif and -Confirm should be allowed
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  Param 	(
    [Parameter(Mandatory=$false,
    HelpMessage='File Path Containing Hotfixes')]
    $FilePath,

    
    [Parameter(Mandatory=$false,
    HelpMessage='Error Log location, default C:\<Command Name>_ErrorLog.txt')]
    [string]$ErrorLog = "c:\$($myinvocation.mycommand)_ErrorLog.txt",
    [switch]$LogErrors
    
  ) #Close Parameters
  
  
  #############################################################
  # Begin Block
  #############################################################
  
  Begin 	{
    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Error log will be $ErrorLog"
    
    # Script Level Variable to Stop Execution if there is an issue with any stage of the script
    $script:EverythingOK = $true
    

    # Hashtable for exit codes

    $exitcodes = @{'0' = 'Success';
      '17301' = 'Error: General Detection error';
      '17302' = 'Error: Applying patch'
      '17303' = 'Error: Extracting file';
      '17021' = 'Error: Creating temp folder';
      '17022' = 'Success: Reboot flag set';
      '17023' = 'Error: User cancelled installation';
      '17024' = 'Error: Creating folder failed';
      '17025' = 'Patch already installed';
      '17026' = 'Patch already installed to admin installation';
      '17027' = 'Installation source requires full file update';
      '17028' = 'No product installed for contained patch';
      '17029' = 'Patch failed to install';
      '17030' = 'Detection: Invalid CIF format';
      '17031' = 'Detection: Invalid baseline';
      '17034' = 'Error: Required patch does not apply to the machine';
      '17038' = 'You do not have sufficient privileges to complete this installation for all users of the machine. Log on as administrator and then retry this installation.';
      '17044' = 'Installer was unable to run detection for this package.';
      '17048' = 'This installation requires Windows Installer 3.1 or greater.';
      }

 
  } #Close Function Begin Block
  
  #############################################################
  # Process Block
  #############################################################
  
  Process {
    
    # First Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        $Installers = Get-ChildItem "$FilePath" -Filter '*.exe'

# Check for Updates before


Foreach ($installer in $installers)
      {
      
      $installerString = $installer.FullName.ToString()

      Write-Verbose $installer.FullName
      $exitcode = (start-Process -FilePath "$installerString" -ArgumentList '/passive' -Wait -Passthru).ExitCode
      $exitcodestring = $null
      $exitcodestring = ($exitcodes."$exitcode")
      Write-Verbose "Exit Code is $exitcode , $exitcodestring "
      Write-Verbose ' '
      }
        
        
      } # Close Try Block
      
      Catch 	{Invoke-ErrorCatchAction} # Close Catch Block
      
      
    } # Close If EverthingOK Block 1
    
    #############################################################
    # Next Script Action or Try,Catch Block
    #############################################################
    
    # Second Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        # Code Goes here
        
        
      } # Close Try Block
      
      Catch 	{Invoke-ErrorCatchAction} # Close Catch Block
      
      
    } # Close If EverthingOK Block 2
    
    
  } #Close Function Process Block 
  
  #############################################################
  # End Block
  #############################################################
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
} #End Function



# Helper Functions below this line ##########################################


    #############################################################
    # Function to Deal with Error Output to Log file
    #############################################################
    

Function Invoke-ErrorCatchAction 
{
  Param 	(
    [Parameter(Mandatory=$false,
    HelpMessage='Switch to Allow Errors to be Caught without setting EverythingOK to False, stopping other aspects of the script running')]
    # By default any errors caught will set $EverythingOK to false causing other parts of the script to be skipped
    [switch]$SetEverythingOKVariabletoTrue
  ) # Close Parameters
  
  # Set EverythingOK to false to avoid running dependant actions
  If ($SetEverythingOKVariabletoTrue) {$script:EverythingOK = $true}
  else {$script:EverythingOK = $false}
  Write-Verbose "EverythingOK set to $script:EverythingOK"
  
  # Write Errors to Screen
  Write-Verbose 'Error Hit'
  Resolve-Error ($Global:Error)[0]
  # If Error Logging is runnning write to Error Log
  
  if ($LogErrors) {
    # Add Date to Error Log File
    Get-Date -format 'dd/MM/yyyy HH:mm' | Out-File $ErrorLog -Append
    $Error | Out-File $ErrorLog -Append
    '## LINE BREAK BETWEEN ERRORS ##' | Out-File $ErrorLog -Append
    Write-Warning "Errors Logged to $ErrorLog"
    # Clear Error Log Variable
    $Error.Clear()
  } #Close If
} # Close Error-CatchActon Function




Function Write-Log{
  
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  param (
    [Parameter(Mandatory=$true,Position=1,HelpMessage='Log String')]
    [string]$LogString,
    
    [string]$LogLevel,
    [string]$LogFilePath
  )
  
  
  #-------------------------------------------------
  # Write-Log
  #-------------------------------------------------
  # Usage:	Writes logging information to file
  # **Parameters are not for interactive execution.**
  # Write-Log -LogString <String> -LogLevel <Error | INFO | Othery> -LogFilePath <LogFilePAth> 
  #
  #-------------------------------------------------
  
  $date = get-date
  
  If($Verbose){write-host $LogString -foregroundcolor green}
  
  $FinalLogString = '['+$date+ ']:'+'['+$LogLevel+']'+':'+$LogString
  
  Add-content $LogFilePath $FinalLogString
  
} # Close Write-Log Function

Function Resolve-Error
{
  
  #############################################################################
  ##
  ## Resolve-Error
  ##
  ## From Windows PowerShell Cookbook (O'Reilly)
  ## by Lee Holmes (http://www.leeholmes.com/guide)
  ##
  ##############################################################################
  
<#

.SYNOPSIS

Displays detailed information about an error and its context.

#>
  
  param(
    ## The error to resolve
    $ErrorRecord = ($error[0])
  )
  
  Set-StrictMode -Off
  
  ''
  'If this is an error in a script you wrote, use the Set-PsBreakpoint cmdlet'
  'to diagnose it.'
  ''
  
  'Error details ($error[0] | Format-List * -Force)'
  '-'*80
  $errorRecord | Format-List * -Force
  
  'Information about the command that caused this error ' +
  '($error[0].InvocationInfo | Format-List *)'
  '-'*80
  $errorRecord.InvocationInfo | Format-List *
  
  'Information about the error''s target ' +
  '($error[0].TargetObject | Format-List *)'
  '-'*80
  $errorRecord.TargetObject | Format-List *
  
  'Exception details ($error[0].Exception | Format-List * -Force)'
  '-'*80
  
  $exception = $errorRecord.Exception
  
  for ($i = 0; $exception; $i++, ($exception = $exception.InnerException))
  {
    "$i" * 80
    $exception | Format-List * -Force
  }
  
} #end function Resolve-Error