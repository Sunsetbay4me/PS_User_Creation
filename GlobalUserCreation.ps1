# Start-Transcript -OutputDirectory C:\Users\Allanl\Desktop\PS_Reports\New_User_Creation_Logs\
########## Change to the correct location ##################
Set-Location C:\Users\Allanl\Documents\Personal\PS
################## Import ActiveDirectory ##################
Import-Module -Name ActiveDirectory
Import-Module -Name NTFSSecurity
##################### Add-Type for GUI #####################
Add-Type -AssemblyName Microsoft.VisualBasic
##################### Add-Type for Voice #####################
Add-Type -AssemblyName System.Speech
    $SpeechSynth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $SpeechSynth.Rate = -2
    $SpeechSynth.SelectVoice('Microsoft Zira Desktop')
    #$SpeechSynth.SelectVoice('Microsoft David desktop')
    #$SpeechSynth.GetInstalledVoices() | Select-Object -ExpandProperty VoiceInfo | Select-Object -Property Culture, Name, Gender, Age
#######################################################################################################################################################################
Clear-Host
#######################################################################################################################################################################
Get-MsolAccountSku
#######################################################################################################################################################################
#
# Set Variables
#
#######################################################################################################################################################################
$BigHashLine = '#########################################################################################################'
#$TempPSWD = 'BWC12345me'
$SNMPServer = '10.10.200.53'
$EnterprisePack = 'brightwood:ENTERPRISEPACK'
$TaskCompleted = 'Task is complete here.'
$MemberOf = 'MemberOf'
$WaitAWhile = 'System will pause for 15 seconds for replication.'
#$VerifyLog = "$env:USERPROFILE\Desktop\PS_Reports\New_User_Creation_Logs\{0}.txt"
$OutPutFile = "$env:SystemDrive\Scripts\Users\Copyfrom.txt"
$OptionA = 'A'
$OptionB = 'B'
$OptionC = 'C'
$OptionD = 'D'
$OptionE = 'E'
$OptionF = 'F'
$OptionG = 'G'
$OptionH = 'H'
$OptionI = 'I'
$OptionJ = 'J'
$OptionK = 'K'
$OptionL = 'L'
$OptionM = 'M'
$OptionN = 'N'
$OptionO = 'O'
$OptionP = 'P'
$OptionQ = 'Q'
$OptionR = 'R'
$OptionS = 'S'
$OptionT = 'T'
$OptionU = 'U'
$OptionV = 'V'
$OptionW = 'W'
$OptionX = 'X'
$OptionY = 'Y'
$OptionZ = 'Z'
$Option1 = '1'
$Option2 = '2'
$Option3 = '3'
$Option4 = '4'
$Option5 = '5'
$Option6 = '6'
$Option7 = '7'
$ExitChoice = '0'
$WIState = 'Wisconsin'
$WIZip_1 = '54751'
$WICity_1 = 'Menomonie'
$NoPoBox = ' '
$HomeDriveLetterG = 'G:'
$HomeDrive_1 = '\\file01\home\'
$StateOR = 'Oregon'
$Error_1 = 'Invalid selection'
$Prompt_1 = '         Type your Choice and press ENTER:              '
$DashLine_1 = '#------------------------------------------------------#'
$HashLine_1 = '#======================================================#'
$Security_1 = 'System.Management.Automation.PSCredential'
$PSSessionsClosed = 'All open PSSEssions have now been closed'
$CheckPSSessions = 'Checking for open PSSessions'
$HashLine = '########################################################'
$Country = 'US'
$AllanL = 'allanl@brightwood.com'
$AD = 'ActiveDirectory'
$Spacer ='  '
$NeedTallyWorks = ''
#Set-Location -Path i:\scripts\users\ -Verbose
################## Import ActiveDirectory ##################
#Import-Module -Name $AD
#Import-Module -Name NTFSSecurity
#######################################################################################################################################################################
Clear-Host
#######################################################################################################################################################################
#
# Set Variables
#
#######################################################################################################################################################################
Write-Output -InputObject 'The variables are being set!!'
$ADAdmin = 'bwc\allanl'
$ADAdminMSOL = $AllanL
$pwd_secure_string = Get-Content -Path "$env:USERPROFILE\documents\windowspowershell\Password.txt" | ConvertTo-SecureString
$AADServer = 'DC01'
$City = $ShortCity = ''
$ok = $ok1 = ''
$PC = $POBox = ''
$PRNumber = ''
$PRNumberPreFix = 'PR#'
$Street = $State = ' '
$Country = $Country
$homedrive = ''
$choice = $choice1 = ''
#$PRNumber = $NewUser = ''
$LoggingScript = ''
$EmpNo = ''
$JobTitle = ''
$Dept = ''
$ShortDept = ''
$Mgr = ''
# $Copy =""
$CopyFrom = $OfficeNew = ''
$Date = (Get-Date -format D)
$Company = 'Bright Wood Corporation'
$session = ''
Write-Output -InputObject 'The variables have been set!!'
Write-Verbose -Message $HashLine
#######################################################################################################################################################################
#
# Run AADSync on DC01
#
#######################################################################################################################################################################
# Run a choice 
Clear-Host
$SpeechSynth.Speak("Do you want to run AADSync")
$Continue = Read-Host "Do you want to run AADSync? (Y/N)?"
while("Y","N" -notcontains $Continue) {
    $Continue = Read-Host "Do you want to run AADSync?  (Y/N)?"
} # WHILE
if ($Continue -eq "Y") {

Write-Output -Verbose -InputObject $CheckPSSessions
Get-PSSession | Remove-PSSession
Write-Output -InputObject $PSSessionsClosed
Start-Sleep -Seconds 15
<#
$seconds = 15
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
write-output -InputObject 'AADSync on DC01 is about to be run.'
#------------------------------------------------------
$Credential = New-Object -TypeName $Security_1 -ArgumentList $ADAdmin, $pwd_secure_string
Invoke-Command -ComputerName $AADServer -Credential $Credential -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
Start-Sleep -Seconds 60
<#
<$seconds = 60
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
Get-PSSession | Remove-PSSession
write-output -InputObject 'AADSync on DC01 has been run.'
#------------------------------------------------------------------
Write-Output "We are all done here:" -ForegroundColor Yellow
#------------------------------------------------------------------
} # IF
#######################################################################################################################################################################
Write-Verbose -Message $HashLine
Write-Output -InputObject 'Time to Make some selections and input data!!'
#######################################################################################################################################################################
#
# Pick a City/Location to associate the employee to
#
#######################################################################################################################################################################
# The menu below allows you to select a location. 
# You then have the option of changing you mind for a new location before entering X to leave the menu.
# Get the new user's Work City.
do{
    do{
        Write-Output $HashLine 
        Write-Output '#              Select A Location                       #' 
        Write-Output $HashLine_1 
        Write-Output '# A = Madras, OR,  B = Redmond, OR, C = Prineville, OR #'
        Write-Output $HashLine_1 
        Write-Output '# D = Menomonie, WI, E = Menomonie North, WI           #'
        Write-Output $HashLine_1              
        Write-Output '# F = Dubuque, IO                                      #' 
        Write-Output $DashLine_1 
        Write-Output '#      0 - Exit to continue after making your choice.  #' 
        Write-Output $HashLine 
        Write-Output $Prompt_1  

        $choice = Read-Host
        # $choice = ($host.UI.RAWUI.ReadKey('NoEcho,IncludeKeyUp')).character
       
        write-output -InputObject ''

        $ok = $choice -match '^[abcdef0]+$'
        

        if (-not $ok) {write-output -InputObject $Error_1}
        } until ($ok)
    switch -Regex ($choice){
        $OptionA {
            write-output -InputObject "You entered 'A' - Madras" 
            $City= 'Madras'
            $ShortCity = 'MAD - '
            $PC= '97741'
            $POBox = 'PO Drawer 828'
            $Street = '335 NW Hess Street'
            $State = $StateOR
           # $Country = "US"
            $homedr1ve = $HomeDrive_1
            $LoggingScript = 'madlogon.vbs'
            $homedriveLetter = $HomeDriveLetterG} # A
        $OptionB {
            write-output -InputObject "You entered 'B' - Redmond" 
            $City= 'Redmond'
            $ShortCity = 'RED - '
            $PC= '97756'
            $POBox = $NoPoBox
            $Street = '630 SE 1st St'
            $State = $StateOR
           # $Country = "US"
            $homedr1ve = '\\red-file01\home\'
            $LoggingScript = 'redlogon.vbs'} # B
        $OptionC {
            write-output -InputObject "You entered 'C' - Prineville" 
            $City= 'Prineville'
            $ShortCity = 'PRV - '
            $PC= '97754'
            $POBox = $NoPoBox
            $Street = '1941 NW Industrial Park Rd'
            $State = $StateOR
           # $Country = "US"
            $homedr1ve = $HomeDrive_1
            $LoggingScript = 'madlogon.vbs'} # C
        $OptionD {
            write-output -InputObject "You entered 'D' - Menomonie" 
            $City= $WICity_1
            $ShortCity = 'MEN - '
            $PC= $WIZip_1
            $POBox = $NoPoBox
            $Street = '6121 Walton Ave'
            $State = $WIState
           # $Country = "US"
            $homedr1ve = $HomeDrive_1
            $homedriveLetter = $HomeDriveLetterG
            $LoggingScript = 'menlogon.vbs'} # D
        $OptionE {
            write-output -InputObject "You entered 'E' - Menomonie North" 
            $City= $WICity_1
            $ShortCity = 'MEN-North - '
            $PC= $WIZip_1
            $POBox = $NoPoBox
            $Street = '5105 Freitag Drive'
            $State = $WIState
           # $Country = "US"
            $homedr1ve = $HomeDrive_1
            $homedriveLetter = $HomeDriveLetterG
            $LoggingScript = 'menlogon.vbs'} # E
        $OptionF {
            write-output -InputObject "You entered 'F' - Dubuque" 
            $City = 'Dubuque'
            $ShortCity = 'DUB - '
            $PC = '52001'
            $POBox = $NoPoBox
            $Street = '1115 Purina Drive Ste 100'
            $State = 'Iowa'
           # $Country = "US"
            $homedr1ve = $HomeDrive_1
            $homedriveLetter = $HomeDriveLetterG
            $LoggingScript = 'menlogon.vbs'} # F
    } # switch
} until ($choice -match $ExitChoice)
#######################################################################################################################################################################
#
# Select a Department to associate the employee to.
#
#######################################################################################################################################################################
Write-Verbose -Message $HashLine
# The menu below allows you to select a location. 
# You then have the option of changing you mind for a new location before entering X to leave the menu.
# Get the new user's Work City.
#--------------------------------------------------------------------------------
Clear-Host
do{
    do{
        Write-Output $HashLine 
        Write-Output '#                                       # Select A Department                   #                                       #' 
        Write-Output $HashLine_1
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              A = Cut                  #              B = FJ                   #              C = Value                #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#' 
        Write-Output '#              D = Maintenance          #              E = Electrical           #              F = PDM/Watch            #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#' 
        Write-Output '#              G = Fabrication          #              H = Grinding Rooms       #              I = Hyster               #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#' 
        Write-Output '#              J = Purchasing/Rebuild   #              K = Shipping             #              L = Accounting           #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              M = Accounts Payable     #              N = Accounts Receivable  #              O = Clerical             #' 
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              P = Costing              #              Q = Drafting             #              R = Fulfillment          #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              S = IT                   #              T = AD                   #              U = DB                   #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              V = Networking           #              W = Support              #              X = Lumber Receiving     #' 
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              Y = Mill Orders          #              Z = Operations           #                                       #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              1 = Personnel            #              2 = Quality              #              3 = Replenishment        #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              4 = Sales                #              5 = Office               #              6 = Paint Line           #'
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output '#              7 = Warehouse            #                                       #                                       #' 
        Write-Output '#-----------------------------------------------------------------------------------------------------------------------#'
        Write-Output $DashLine_1 
        Write-Output '#      0 - Exit to continue after making your choice.  #' 
        Write-Output $HashLine 
        Write-Output $Prompt_1 

        $choice1 = Read-Host
        # $choice = ($host.UI.RAWUI.ReadKey('NoEcho,IncludeKeyUp')).character
       
        write-output -InputObject ''

        $ok1 = $choice1 -match '^[abcdefghijklmnopqrstuvwxyz12345670]+$'
        

        if (-not $ok1) {write-output -InputObject $Error_1 }
        } until ($ok1)
    switch -Regex ($choice1){
        $OptionA {write-output -InputObject "You entered 'A' - Cut" 
            $ShortDept = "Cut"
            $NeedTallyWorks = "Please assign rights TallyWorks for this user."} # A
        $OptionB {write-output -InputObject "You entered 'B' - FJ" 
            $ShortDept = 'FJ'} # B
        $OptionC {write-output -InputObject "You entered 'C' - Value" 
            $ShortDept = 'Value'} # C
        $OptionD {write-output -InputObject "You entered 'D' - Maintenance" 
            $ShortDept = 'MNT'
            $NeedMaximo = "Please assign rights Maximo for this user."} # D
        $OptionE {write-output -InputObject "You entered 'E' - Electrical" 
            $ShortDept = 'ELECT'
            $NeedMaximo = "Please assign rights Maximo for this user."} # E
        $OptionF {write-output -InputObject "You entered 'F' - PDM/Watch" 
            $ShortDept = 'PDM'} # F
        $OptionG {write-output -InputObject "You entered 'G' - Fabrication"
            $ShortDept = 'FAB'
            $NeedMaximo = "May need assign rights Maximo for this user."} # G
        $OptionH {write-output -InputObject "You entered 'H' - Grinding Rooms" 
            $ShortDept = 'GRIND'
            $NeedMaximo = "Please assign rights Maximo for this user."} # H
        $OptionI {write-output -InputObject "You entered 'I' - Hyster" 
            $ShortDept = 'HYSTER'
            $NeedMaximo = "Please assign rights Maximo for this user."} # I
        $OptionJ {write-output -InputObject "You entered 'J' - Purchasing" 
            $ShortDept = 'PURCH'
            $NeedMaximo = "Please assign rights Maximo for this user."} # J
        $OptionK {write-output -InputObject "You entered 'K' - Shipping" 
            $ShortDept = 'SHIP'} # K
        $OptionL {write-output -InputObject "You entered 'L' - Accounting" 
            $ShortDept = 'ACCT'
            $LoggingScript = 'AcctLogon.vbs'} # L
        $OptionM {write-output -InputObject "You entered 'M' - Accounts Payable"
            $ShortDept = 'A/P'
            $LoggingScript = 'AcctLogon.vbs'} # M
        $OptionN {write-output -InputObject "You entered 'N' - Accounts Receivable" 
            $ShortDept = 'A/R'
            $LoggingScript = 'AcctLogon.vbs'} # N
        $OptionO {write-output -InputObject "You entered 'O' - Clerical" 
            $ShortDept = 'Clerical'} # O
        $OptionP {write-output -InputObject "You entered 'P' - Costing" 
            $ShortDept = 'Costing'} # P
        $OptionQ {write-output -InputObject "You entered 'Q' - Drafting" 
            $ShortDept = 'Drafting'} # Q
        $OptionR {write-output -InputObject "You entered 'R' - Fullfillment" 
            $ShortDept = 'Fullfillment'} # R
        $OptionS {write-output -InputObject "You entered 'S' - IT" 
            $ShortDept = 'IT'} # S
        $OptionT {write-output -InputObject "You entered 'T' - Application Development" 
            $ShortDept = 'IT-AD'} # T
        $OptionU {write-output -InputObject "You entered 'U' - DataBase" 
            $ShortDept = 'IT-DB'} # U
        $OptionV {write-output -InputObject "You entered 'V' - Networking" 
            $ShortDept = 'IT-Networking'} # V
        $OptionW {write-output -InputObject "You entered 'W' - Support" 
            $ShortDept = 'IT-Support'
            $LoggingScript = 'MadLogonNo7.vbs'} # W
        $OptionX {write-output -InputObject "You entered 'X' - Lumber Receiving" 
            $ShortDept = 'Lumber Receiving'
            $NeedTallyWorks = "Please assign rights TallyWorks for this user."} # X
        $OptionY {write-output -InputObject "You entered 'Y' - Mill Orders" 
            $ShortDept = 'Mill Orders'} # Y
        $OptionZ {write-output -InputObject "You entered 'Z' - Operations" 
            $ShortDept = 'Operations'} # Z
        $Option1 {write-output -InputObject "You entered '1' - Personnel" 
            $ShortDept = 'Personnel'} # 1
        $Option2 {write-output -InputObject "You entered '2' - Quality" 
            $ShortDept = 'Quality'} # 2
        $Option3 {write-output -InputObject "You entered '3' - Replenishment" 
            $ShortDept = 'Replenishment'} # 3
        $Option4 {write-output -InputObject "You entered '4' - Sales" 
            $ShortDept = " Sales"} # 4
        $Option5 {write-output -InputObject "You entered '5' - Office" 
            $ShortDept = " Office"} # 5
        $Option6 {write-output -InputObject "You entered '6' - PaintLine" 
            $ShortDept = " Paint-Line"} # 6
        $Option7 {write-output -InputObject "You entered '7' - Warehouse" 
            $ShortDept = " Warehouse"} # 7
    }
} until ($choice1 -match $ExitChoice)
#######################################################################################################################################################################
Clear-Host
Write-Verbose -Message $HashLine
#######################################################################################################################################################################
#
# Gets all of the required info to be copied to the new account
#
#######################################################################################################################################################################
# enter the HD/PR Number e.g 12345
$SpeechSynth.Speak("Please Enter the PR Number")
$PRNumber = [Microsoft.VisualBasic.Interaction]::InputBox("Enter The PR Number without PR#:" , "PR Number")
$PRNumber = $PRNumberPrefix + $PRNumber
#--------------------------------------------------------------------------------
# enter login name of the user user to copy memberships from.
#Add-Type -AssemblyName Microsoft.VisualBasic
$SpeechSynth.Speak("Please Enter the User Name to Copy from")
$CopyFrom = [Microsoft.VisualBasic.Interaction]::InputBox("Enter username to copy memberships from: " , "Copy From")
$CopyFromName = Get-ADUser -Identity $CopyFrom -Properties Name, DisplayName, Department, Office
$CopyFromName1 = $CopyFromName.Name
# $CopyFromDepartment = $CopyFromName.Department
$CopyFromOffice = $CopyFromName.Office
#$CopyFromOffice
#--------------------------------------------------------------------------------
Get-ADUser -Identity $CopyFrom -Properties Office, Department, ScriptPath, Description, Title, ExtensionAttribute1, Manager | Out-File -FilePath $OutPutFile
Start-Sleep -Seconds 15
<#
$seconds = 15
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
& "$env:windir\system32\notepad.exe" $OutPutFile
#--------------------------------------------------------------------------------
# enter login name of the new user.
$SpeechSynth.Speak("Please Enter the user's Name")
$NewUser = [Microsoft.VisualBasic.Interaction]::InputBox("Enter User's Name" , "Name")
#--------------------------------------------------------------------------------
# enter employeeid for the new user.
$SpeechSynth.Speak("Please Enter the Employee's ID Number")
$EmpNo = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Employee's ID: " , "Employee ID #")
#--------------------------------------------------------------------------------
# enter Office for the new user.
#$OfficeNew = read-host -Prompt "Enter Employee's Office: "
#$OfficeNew = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Employee's Office: " , "Employee Office")
$OfficeNew = $CopyFromOffice #= $CopyFrom.Office
#--------------------------------------------------------------------------------
# enter Logon Script for the new user.
# $LoggingScript = read-host -Prompt "Enter Employee's Logon Script: "
#--------------------------------------------------------------------------------
# enter new employee's job title.
$SpeechSynth.Speak("Please Enter the users Job Title")
$JobTitle = [Microsoft.VisualBasic.Interaction]::InputBox("Enter User's Job Title" , "Job Title")
#--------------------------------------------------------------------------------
# enter the new employee's Manager's name.
$SpeechSynth.Speak("Please Enter the user's Manager's Name")
$Mgr = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Manager's Name" , "Manager's Name")
$MgrName = Get-ADUser -Identity $Mgr -Properties Name, DisplayName
$MgrName1 = $MgrName.Name
#--------------------------------------------------------------------------------
Write-Verbose -Message $HashLine
#######################################################################################################################################################################
#
# Build User creation notes DO NOT EDIT
#
#$NewUser = "allanl" # For testing Only
#######################################################################################################################################################################
$NewUserName = Get-ADUser -Identity $NewUser -Properties *
#$NewUserName # For Testing ONLY
$NewUserName1 = $NewUserName.name
$info1 ='Created as per '
$info2 = $PRNumber
$info3 = ' by Allan Laird on '
$info4 = $date
$newnotes = $info1 +$info2 +$info3 + $info4
$emaildomain = '@brightwood.com'
$homepage = 'http://brightwood.com'
$homedrive = $homedr1ve + $NewUser
$Dept = $ShortCity + $ShortDept
$UPN = $NewUser + $emaildomain
$VerifyLog = "$env:USERPROFILE\Desktop\PS_Reports\New_User_Creation_Logs\{0}.txt"
$UserCreationLog = $info1 + $newnotes + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser)
$UserCreationLog = 'UserName: ' + $NewUser + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
$UserCreationLog = 'DisplayNameName: ' + $NewUserName1 + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
$UserCreationLog = 'EmployeeID: ' + $EmpNo + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
#$UserCreationLog = 'TempPassword: ' + $TempPSWD + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append  
$UserCreationLog = 'Office: ' + $OfficeNew + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
$UserCreationLog = 'Login Script: ' + $LoggingScript + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
$UserCreationLog = 'Title: ' + $JobTitle + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
$UserCreationLog = 'Manager: ' + $MgrName1 + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser)  -Append
$UserCreationLog = 'HomeDrive: ' + $homedrive + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser)  -Append
$UserCreationLog = 'Dept: ' + $Dept + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser)  -Append
$UserCreationLog = 'EmailAddress: ' + $UPN + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser)  -Append
Write-Verbose -Message $HashLine
#######################################################################################################################################################################
#
# The heavy Lifting Starts Here
#
#######################################################################################################################################################################
# Write-Output -InputObject ' The magic begins!'
$SpeechSynth.Speak("heavy Lifting is about to Start")
Write-Output -InputObject 'The Heavy Lifting Starts Here!'
Write-Verbose -Message $HashLine
#######################################################################################################################################################################
#
# Set new User's attributes
#
#######################################################################################################################################################################
Set-ADUser -Identity $NewUser -EmployeeID $EmpNo -StreetAddress $Street -City $City -State $State -PostalCode $PC -Country $Country -Company $Company -POBox $POBox -Title $JobTitle -Office $OfficeNew -Department $Dept -ScriptPath $LoggingScript -Manager $Mgr -HomePage $homepage -Replace @{info=$newnotes} -Verbose
$SpeechSynth.Speak("Updating New User")
write-output -InputObject 'New User has been successfully Updated.', $newnotes
$UserCreationLog + ('Group Memberships will be copied from: {0} ' -f $CopyFromName1) + $Spacer | Out-file -FilePath ($VerifyLog -f $NewUser) -Append
$UserCreationLog = 'Login Script: ' + $LoggingScript + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
# $UserCreationLog = 'Office: ' + $OfficeNew + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
write-output -InputObject $WaitAWhile
Start-Sleep -Seconds 15
<#
$seconds = 15
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
Write-Verbose -Message $HashLine
#######################################################################################################################################################################
# copy-paste process. Get-ADUser membership     | then select memberships                       | and add them to the second user
#######################################################################################################################################################################
Get-ADUser -identity $CopyFrom -Properties $MemberOf | Select-Object -Property $MemberOf -ExpandProperty $MemberOf | Add-ADGroupMember -Members $NewUser
$SpeechSynth.Speak("Copying User's Rights")
write-output -InputObject "New User's Group Memberships are being copied from ", $CopyFrom"."
write-output -InputObject $WaitAWhile
Start-Sleep -Seconds 15
<#
$seconds = 15
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
$UserCreationLog + ('Group Memberships successfully copied from: {0} ' -f $CopyFromName1) + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
Write-Verbose -Message $HashLine
#######################################################################################################################################################################
# Create the user's Home Directory.
#######################################################################################################################################################################
$SpeechSynth.Speak("Creating User's Home Drive")
New-Item -path $homedrive -ItemType Directory -Force
write-output -InputObject "New User's Home Drive successfully created."
Add-NTFSAccess -Path $homedrive -Account $NewUser -AccessRights FullControl -AppliesTo ThisFolderSubfoldersAndFiles -PassThru
Get-NTFSAccess -Path $homedrive
# write-output -InputObject 'Check New User access rights in Security tab in Home drive Properties Folder.' 
write-output -InputObject 'System will pause for 30 seconds for replication.'
Start-Sleep -Seconds 30
<#
$seconds = 30
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
$UserCreationLog + 'HomeDrive has been created and rights assigned:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
#######################################################################################################################################################################
# Create a Lync/Skype for Business account
#######################################################################################################################################################################
Clear-Host
Write-Verbose -Message $HashLine
Write-Output -Verbose -InputObject $CheckPSSessions
Get-PSSession | Remove-PSSession
Write-Output -InputObject $PSSessionsClosed
Start-Sleep -Seconds 15
<#
$seconds = 15
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
$SpeechSynth.Speak("Creating User's Lync Account")
write-output -InputObject 'LYNC Account Creation is about to start'
$Credential = New-Object -TypeName $Security_1 -ArgumentList $ADAdmin, $pwd_secure_string
$session = New-PSSession -ConnectionUri 'https://lync01.bwc.brightwood.com/OcsPowershell' -Credential $credential
#--------------------------------------------------------------------------------
Import-PSSession -Session $session -AllowClobber
Import-Module -Name $AD
Start-Sleep -Seconds 45
Enable-CsUser -Identity $NewUser -RegistrarPool 'lync01.bwc.brightwood.com' -SipAddressType SamAccountName -SipDomain brightwood.com
Start-Sleep -Seconds 15
Grant-CsClientPolicy -Identity $NewUser -PolicyName ForceADPictures
Get-PSSession | Remove-PSSession 
write-output -InputObject "New User's Lync Account successfully created."
Write-Output -InputObject 'A Client Policy has been added to the Lync account.'
$UserCreationLog + 'Lync Account Has been created:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append 
write-output -InputObject $TaskCompleted
#######################################################################################################################################################################
#
# Add User's picture to AD if one exists.
#
#######################################################################################################################################################################
#$EmpNo = "16465"
Write-Verbose -Message $HashLine
write-output -InputObject "Checking for user's picture:"
$AddOn = "*"
$Named = $AddOn + $EmpNo + $AddOn
#$Named
Get-ChildItem -Path "h:\employee pics\adp\" -Include $Named -file -Recurse -ErrorAction SilentlyContinue | Format-Table -AutoSize Name
$SpeechSynth.Speak("if the user has an AD picture you can add it now. If not request one be added by Personnel") 
#$UserPicture = Get-ChildItem -Path "h:\employee pics\adp\" -Include $Named -file -Recurse -ErrorAction SilentlyContinue #| Format-Table -AutoSize Name
#$UserPicture

#Set-UserPhoto -Identity $NewUsers -PictureData ([System.IO.File]::ReadAllBytes("h:\employee pics\adp\$Named")) -WhatIf
write-output -InputObject "Add the user's picture into AD. System Will Pause "
#Set-ADUser -Identity $NewUser -thumbnailPhoto=([byte[]](Get-Content "\\file01\MultiApp\Employee Pics\ADP\$pictureID1" -Encoding byte))
#Set-ADUser $NewUser -Replace @{thumbnailPhoto=$photo}
Pause
$UserCreationLog + 'Users picture (if one exists) has been added to User Object:' + $Spacer | Out-file -FilePath ($VerifyLog -f $NewUser) -Append
#######################################################################################################################################################################
#
# Assign an Office 365 License to the employee
#
#######################################################################################################################################################################
$Security_1 = 'System.Management.Automation.PSCredential'
$HashLine = '########################################################'
$AllanL = 'allanl@brightwood.com'
$ADAdminMSOL = $AllanL
$pwd_secure_string = Get-Content -Path "$env:USERPROFILE\documents\windowspowershell\Password.txt" | ConvertTo-SecureString
#--------------------------------------------------------------------------------
Write-Verbose -Message $HashLine
$Credential = New-Object -TypeName $Security_1 -ArgumentList $ADAdminMSOL, $pwd_secure_string
Get-MsolUser -UserPrincipalName $UPN | Format-Table -Property UserPrincipalName, DisplayName, isLicensed, UsageLocation
write-output -InputObject 'Check that this the user you created.  System Will Pause'
Pause
#--------------------------------------------------------------------------------
$SpeechSynth.Speak("Creating Users Office three sixty five License")
$Myo365SkuOption = New-MsolLicenseOptions -AccountSkuId $EnterprisePack -DisabledPlans 'YAMMER_ENTERPRISE','POWERAPPS_O365_P2'
Set-MsolUser -UserPrincipalName $UPN -UsageLocation $Country -Verbose
Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $EnterprisePack -LicenseOptions $Myo365SkuOption -Verbose
Start-Sleep -Seconds 15
<#
$seconds = 15
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
Get-MsolUser -UserPrincipalName $UPN | Format-Table -Property UserPrincipalName, DisplayName, isLicensed, UsageLocation
Start-Sleep -Seconds 30
<#
$seconds = 30
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
$UserCreationLog + 'Office 365 License has been assigned:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
# Get-Mailbox –identity $UPN | Set-Clutter -enable $false
# $UserCreationLog + 'Clutter has been disbled:' + $Spacer | Out-file -FilePath ($VerifyLog -f $NewUser) -Append
write-output -InputObject $TaskCompleted
#######################################################################################################################################################################
# Clear-Host
# write-output "Looking to see is there any Shared Mailboxes that will need to be associated with the new user. PLEASE WAIT." 
# Get-Recipient -RecipientTypeDetails SharedMailbox | Get-MailboxPermission -User $CopyFrom | Format-Table -AutoSize
# Start-Sleep -Seconds 45
#######################################################################################################################################################################
#
# Add any other login rights and or phone extension updates, then send an email to Personnel to update EE-Data
#
#######################################################################################################################################################################
$SpeechSynth.Speak("Time to tidy up the finer details")
Write-Output $NeedTallyWorks 
Write-Output $NeedMaximo 
do{
    do{
        write-output -InputObject $HashLine 
        write-output -InputObject '#                    Select A Task                     #' 
        write-output -InputObject $HashLine_1 
        write-output -InputObject '#              A = Maximo login needed                 #' 
        write-output -InputObject '#              B = Adagio login needed                 #' 
        write-output -InputObject '#              C = Lube-IT login needed                #'
        write-output -InputObject '#              D = Plant Access in POM assignment      #'
        write-output -InputObject '#              E = Phone List Update                   #' 
        write-output -InputObject '#              F = BWC Reports Access Copied           #'
        write-output -InputObject '#              G = TallyWorks Login Request            #'
        write-output -InputObject '#              H = KMNet Login Requested               #'
        write-output -InputObject '#              I = Not Used Yet                        #' 
        write-output -InputObject $DashLine_1 
        write-output -InputObject '#         0 - Exit to continue when finished.          #' 
        write-output -InputObject $HashLine 
        write-output -InputObject $Prompt_1 

        $choice = Read-Host
        # $choice = ($host.UI.RAWUI.ReadKey('NoEcho,IncludeKeyUp')).character
       
        write-output -InputObject ''

        $ok = $choice -match '^[abcdefgh0]+$'
        

        if (-not $ok) {write-output -InputObject $Error_1}
        } until ($ok)
    switch -Regex ($choice){
        $OptionA {
            $Info = Get-ADUser -Identity $NewUser
            $Info= $Info.Name 
            Write-Verbose -Message $HashLine
            write-output -InputObject 'Notifing Troy Towers for any Maximo access that is required, he will also need to know who to give access like.'
            Send-MailMessage -From $AllanL -To UG_PurchMgr@brightwood.com -Subject 'Maximo Account login needed' -Body "We have a new user $NewUserName1 with the login name $NewUser in the system and they need a Maximo login as per $PRNumber same rights as $CopyFrom . Please send details to $MgrName1 ." -SmtpServer $SNMPServer
            Write-Output -InputObject ' Email has been sent to Troy Towers'
            $UserCreationLog + 'Maximo login has been requested:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append} # A
        $OptionB {
            write-output -InputObject $BigHashLine 
            write-output -InputObject 'Notifing Davis Barringer for all Adagio Access requests.'
            Send-MailMessage -From $AllanL -To UG_NetworkMgr@brightwood.com -Subject 'Adagio Account login needed' -Body "We have a new user $NewUserName1 in the system and they need an Adagio login as per $PRNumber . Thanks." -SmtpServer $SNMPServer
            Write-Output -InputObject 'Email has been sent to Davis Barringer'
            $UserCreationLog + 'Adagio login has been requested:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append} # B
        $OptionC {
            write-output -InputObject $BigHashLine 
            write-output -InputObject 'Notifing Jesse King for Lube-It access requests.'
            Send-MailMessage -From $AllanL -To JesseK@brightwood.com -Subject 'Lube-IT Account login needed' -Body "We have a new user $NewUserName1 in the system and they need a Lube-It login as per $PRNumber . Thanks." -SmtpServer $SNMPServer
            Write-Output -InputObject ' Email has been sent to Jesse King'
            $UserCreationLog + 'Lube-IT login has been requested:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append} # C
        $OptionD {
            write-output -InputObject $BigHashLine 
            write-output -InputObject 'Set Plant Access in POM with Pro Schedule.'
            #Pause
            $PlantAccess = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Plants numbers access is required for: " , "Plant Access #")
            #$PlantAccess = Read-Host  'Enter the Plant Number(s): ' 
            Write-Output -InputObject ' Access to Plant(s) has been assigned ' $PlantAccess
            $UserCreationLog + 'POM rights has been requested to Plants: ' + $PlantAccess + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append} # D
        $OptionE {
            write-output -InputObject $BigHashLine 
            write-output -InputObject 'Notify Clerical for any Phone List Extension Changes.'
            $ExtNumber = Read-Host -Prompt 'Enter the new extension number: '
            Send-MailMessage -From $AllanL -To UG_ClericalMgr@brightwood.com -Subject 'Phone List Update' -Body "We have a new user $NewUserName1 in the system and they will be at Extension $ExtNumber . Thanks."  -SmtpServer $SNMPServer
            Write-Output -InputObject ' Email has been sent to Clerical Dept'
            $UserCreationLog + 'Phone List update has been requested:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append} # E
        $OptionF {
            write-output -InputObject $BigHashLine
            Write-Output "Assign BWC Reports to: $NewUser"
            $UserCreationLog + 'BWCReports will be assigned with the same rights as ' + $CopyFromName1 + $Spacer | Out-file -FilePath ($VerifyLog -f $NewUser) -Append} # F
        $OptionG {
            write-output -InputObject $BigHashLine
            write-output -InputObject 'New User needs Tallyworks login.'
            Write-Output -InputObject ' Sending email to IT Support Dept'
            Send-MailMessage -From $AllanL -To UG_ITSupport@brightwood.com -Subject 'New User needs Tallyworks login.' -Body "We have a new user $NewUser in the system and they need a TallyWorks login." -SmtpServer $SNMPServer
            Write-Output -InputObject 'Email has been sent to IT Support Dept'
            $UserCreationLog + 'TallyWorks login has been requested:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append} # G
       $OptionH {
            write-output -InputObject $BigHashLine
            write-output -InputObject 'New IT Member needs KMNet login.'
            Write-Output -InputObject 'Sending email to IT Support Dept'
            Send-MailMessage -From $AllanL -To UG_ITSupport@brightwood.com -Subject 'New User needs KMNet login.' -Body "We have a new user $NewUser in the system and they need a KMNet login." -SmtpServer $SNMPServer
            Write-Output -InputObject ' Email has been sent to IT Support Dept'
            $UserCreationLog + 'KMNet login has been requested:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append} # H        
        $OptionI {
            } # I
    } # switch
} until ($choice -match $ExitChoice)
###################################################################################################################################################################
#######################################################################################################################################################################
$SpeechSynth.Speak("Do you to run EEData Update")
$Continue =""
$Continue = Read-Host "Do you want to run EEData Update"
while("Y","N" -notcontains $Continue) {
    $Continue = Read-Host "Do you want to run E-Data Update Yet?  (Y/N)?"
} # WHILE
if ($Continue -eq "Y") {
write-output -InputObject $BigHashLine
            #write-output -InputObject 'Contacting Personnel and have them run E-Data to enter the user into COM-Individual.'
            Write-Output -InputObject ' System will pause to for you to run EEData.'
            Pause
            #Send-MailMessage -From $AllanL -To UG_Personnel@brightwood.com -Subject 'Please run E-Data/Pers-Tran at earliest opportunity.' -Body "We have a new user $NewUserName1 in the system and data needs to be updated. Thanks." -SmtpServer $SNMPServer
            Write-Output -InputObject ' EEData has been run'
            $UserCreationLog + 'Ee-Data update has been run:' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
            }
if ($Continue -eq "N") {
write-output -InputObject $BigHashLine
            #write-output -InputObject 'Contacting Personnel and have them run E-Data to enter the user into COM-Individual.'
            Write-Output -InputObject ' System will not pause to for you to run EEData.'
            #Pause
            #Send-MailMessage -From $AllanL -To UG_Personnel@brightwood.com -Subject 'Please run E-Data/Pers-Tran at earliest opportunity.' -Body "We have a new user $NewUserName1 in the system and data needs to be updated. Thanks." -SmtpServer $SNMPServer
            Write-Output -InputObject ' EEData will be run be run after the last of the batch has been created'
            $UserCreationLog + 'EEData will be run be run after the last of this batch of new users has been created' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
            }
###################################################################################################################################################################
#######################################################################################################################################################################
$Mgr = $Mgr + $emaildomain
write-output -InputObject $BigHashLine
            #write-output -InputObject 'Contacting relevant managerto advise them the new user will have to watch the Email Security Video.'
            Write-Output -InputObject " Sending email to User's Manager and IT Support. User needs to watch Email Security Video before account will be enabled and password isued."
            Send-MailMessage -From $AllanL -To "$Mgr","UG_ITSupport@brightwood.com" -Subject "$PRNumber :  Email Security Training Video must to be viewed." -Body "We have a new user $NewUserName1 in the system and they are required to watch the Email Security Training Video. The Attendance Sheet needs to be signed by the new user, and a copy sent to IT Support before a password will be assigned and the account the New User account enabled. Thanks." -SmtpServer $SNMPServer
            Write-Output -InputObject "Email has been sent to the User's Manager, and IT Support."
            $UserCreationLog + 'Email Security Training has been requested.' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append

#######################################################################################################################################################################
#
# Tidy up the loose ends and close the script
Set-ADUser -Identity $NewUser -Enabled $false
$UserCreationLog + '## NOTICE Account is Disabled until ALL pre-requisites are completed:##' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
$UserCreationLog + 'BWC Reports will be assigned after Ee-Data has been updated.' | Out-file -FilePath ($VerifyLog -f $NewUser) -Append
$UserCreationLog + 'Password will be assigned after the Attendance Sheet for Email Security Training has been received.:' | Out-file -FilePath ($VerifyLog -f $NewUser) -Append
#
#######################################################################################################################################################################
# write-output -InputObject 'After you have completed all the above, please hit ENTER to continue.'
Pause
$UserCreationLog + '################### END OF STAGE 1 ####################' + $Spacer| Out-file -FilePath ($VerifyLog -f $NewUser) -Append
Stop-Process -name notepad -Force
Start-Sleep -Seconds 5
<#
$seconds = 5
1..$seconds |
ForEach-Object { $percent = $_ * 100 / $seconds; 
 
  Write-Progress -Activity Break -Status "$($seconds - $_) seconds remaining..." -PercentComplete $percent; 
  
  Start-Sleep -Seconds 1
  } 


#>
Remove-Item -Path $OutPutFile
Remove-Variable -Name pwd_secure_string -Verbose
# Pause
write-output -InputObject =, 'You CANNOT close the PR yet.'
$SpeechSynth.Speak("Finished Part 1 of the process.")
#######################################################################################################################################################################
# Stop-Transcript