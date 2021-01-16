# Author : Greg Nimmo
# Version : 0.10
# Description : script to search a users mailbox using known credentials for specified keywords
# todo : any emails that are unread do not change them to read add either support for MFA or use of browser session authentication for a workaround to MFA

# load the exchange dll you can download from https://www.microsoft.com/en-us/download/details.aspx?id=42951
$exchangeDll = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

# check the exchagne dll is available and load it
$checkDll = Test-Path -Path $exchangeDll
if ($checkDll){
    [void][Reflection.Assembly]::LoadFile($exchangeDll)
    Write-Host('Exchange Web Services DLL loaded successfully')
    } else{
    Write-Host('Exchange Web Services DLL not found, check you have installed the DLL')
    }

# main program 
$selection = $true
while ($selection){

    # gather credentials needs to be changed to prompt for any credentials (see note below about secure string issue)
    $username = read-host('Username ')
    $password = Read-Host('Password ') -AsSecureString

    #convert the password into plaintext for use in office 365 connection
    $bitStream = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
    $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bitStream)

    # create a EWS service object
    $service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService 

    # add credentials to the service object
    $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($username,$plainText)
    
    # this TestUrlCallback is purely a security check
    $TestUrlCallback = {
        param ([string] $url)
        if ($url -eq "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml") {
        $true} else {$false}
        }
    
    # Autodiscover using the username (email address) set above
    $service.AutodiscoverUrl($username,$TestUrlCallback)

    # request a text file to use for keyword searching
    do {
        # keyword files are loaded from /user/<username>/Documents directory
        $filePath = "$HOME\Documents\"
        $searchFile = Read-Host -Prompt 'Enter filename (including extension) containing keywords '
        $checkFile = Test-Path -Path "$filePath\$searchFile"
        cd $filePath # changes this so read and write of files uses absolute path
        # write results to the following file
        $MailHarvestResults = $filePath + "OutlookScraperResults_" + (Get-Date -Format dd-MM-yyyy) + ".txt"

        if ($checkFile){
            write-host('Keyword file loaded successfully')
            write-host('Performing keyword search, please wait...')
                # create a folder object 
                $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000) # maximum number of folders to retreive
                $folderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.Webservices.Data.BasePropertySet]::FirstClassProperties)
                $folderView.PropertySet.Add([Microsoft.Exchange.Webservices.Data.FolderSchema]::DisplayName)
                $folderView.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep # recursivly search the folders

                # gather a list of folders
                $folderList = $service.FindFolders([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot, $folderView)

                # loop through each folder found in the users directory
                foreach ($folder in $folderList.Folders){
                    $folderEmails = $folder.FindItems(1000) # maximum number of objects to find in each folder
                    # for each object (email) list its contents
                    foreach ($email in $folderEmails){
                        # loop through each line in the specified text file
                        foreach ($line in (Get-Content $searchFile)) {
                            if ($email.Subject -match ".*$line.*" -or $email.Body.Text -match ".*$line.*"){
                                # Output the results
                                $from = "From: $($email.From.Name)"
                                $subject = "Subject: $($email.Subject)"
                                $body = "Body: $($email.Body.Text)"
                                Write-Output $from $subject $body | Out-File -Append $MailHarvestResults
                                
                            }
                            else{
                                continue 
                            }
                        }
                    }
                }

            }
            else{
                write-host('File not found')
            }
        } while (-not $checkFile)

        # inform user of output
        Write-Host("Search results written to $MailHarvestResults")

        # Ask user to perform another search
        $newSearch = $true
        do{
            $choice = Read-Host('Perform another search (Y/N) ')
            if ($choice.ToUpper() -eq 'Y'){
                $newSearch = $false
                }
            elseif ($choice.ToUpper() -eq 'N'){
                # end the program
                $newSearch = $false
                $selection = $false
            }
            else{
                write-host('Invalid option')
            }
        }while ($newSearch -eq $true)

}

    
