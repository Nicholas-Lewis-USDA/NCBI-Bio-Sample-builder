<#
    .SYNOPSIS
        Creates an Excel sheet containing sample name, SRR #, Bio Project, and who submitted the sample by searching NCBI
    .DESCRIPTION
        User must supply either the sample name or SRR and allows for copy and paste from a table with a single column
        Known bugs -> must be ran in PowerShell - not the editor as the editor doesn't read table rows
    .EXAMPLE
        Output from a sucessfull run
            Enter data, must be one per line and blank line will start excel program: SRR1791702
            Enter data, must be one per line and blank line will start excel program: SRR1791703
            Enter data, must be one per line and blank line will start excel program: SRR1791715
            Enter data, must be one per line and blank line will start excel program: SRR1791735
            Enter data, must be one per line and blank line will start excel program: SRR1791736
            Enter data, must be one per line and blank line will start excel program: SRR1791737
            Enter data, must be one per line and blank line will start excel program: SRR1791752
            Enter data, must be one per line and blank line will start excel program: SRR1791755
            Enter data, must be one per line and blank line will start excel program: SRR1791758
            Enter data, must be one per line and blank line will start excel program: 02-0011
            Enter data, must be one per line and blank line will start excel program: 02-0357
            Enter data, must be one per line and blank line will start excel program: 02-0493
            Enter data, must be one per line and blank line will start excel program: 02-0585
            Enter data, must be one per line and blank line will start excel program: 02-1331
            Enter data, must be one per line and blank line will start excel program: 02-1368
            Enter data, must be one per line and blank line will start excel program: 02-1405
            Enter data, must be one per line and blank line will start excel program: 02-1463
            Enter data, must be one per line and blank line will start excel program: 02-1531
            Enter data, must be one per line and blank line will start excel program: 02-1579
            Enter data, must be one per line and blank line will start excel program:
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791702
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300060
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791703
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300061
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791715
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300073
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791735
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300093
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791736
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300094
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791737
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300095
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791752
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300110
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791755
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300113
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791758
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300116
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-0011
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-0357
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-0493
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-0585
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1331
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1368
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1405
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1463
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1531
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1579
            Writing srr.xlsx...
            Program has completed, you may now exit this window by pressing enter closing the screen.:

        Output file reuslts ->
            Sample_name	SRR #	    Bio Project	Submitted By	URL
            01-0467 	SRR1791702	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791702
            01-0843 	SRR1791703	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791703
            01-2374 	SRR1791715	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791715
            01-4106 	SRR1791735	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791735
            01-4280 	SRR1791736	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791736
            01-4283 	SRR1791737	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791737
            01-5745 	SRR1791752	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791752
            01-6106 	SRR1791755	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791755
            01-6318 	SRR1791758	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791758
            02-0011	    Not Found	Not Found	Not Found	                                                                                                                            https://www.ncbi.nlm.nih.gov/sra/?term=02-0011
            02-0357	    SRR1791767	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-0357
            02-0493	    SRR1791770	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-0493
            02-0585	    SRR1791773	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-0585
            02-1331	    SRR1791777	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-1331
            02-1368	    SRR1791778	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-1368
            02-1405	    SRR1791781	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-1405
            02-1463	    SRR1791782	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-1463
            02-1531	    SRR1791785	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-1531
            02-1579	    SRR1791788	PRJNA251692	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	https://www.ncbi.nlm.nih.gov/sra/?term=02-1579

#>

#Imports and Frame Setup
Add-Type -AssemblyName PresentationFramework
[Net.ServicePointManager]::SecurityProtocol =[Net.SecurityProtocolType]::Tls12

#Gets local user - Windows tested only
function get_user() {
    <#
        .SYNOPSIS
            Gets the current window user
        
        .OUTPUTS
            returns the current window user as a String
    #>
    $local_name = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name #gets local user name
    $local_name = $local_name -split "\\"
    $local_name = $local_name[1]
    
    return $local_name
}

#gets path info for output, assumes windows and will place on Desktop, $user is required -> use get_user() to pull this info
function get_local_path($user) {
     <#
        .SYNOPSIS
            Gets the the path to the user's Desktop
        
        .INPUTS
            Must supply the $user -> get_user will result in this required input
        
        .OUTPUTS
            Returns the path to the user's Desktop as a String
    #>
    $local_path = Join-Path -Path "C:\Users\" -Childpath $user
    $local_path = Join-Path -Path $local_path -Childpath "Desktop"
    return $local_path
}

#Prevent overwriting file on startup $outputpath is required -> use get_local_path() to pull this info
function search_prev_file($local_path) {
    <#
        .SYNOPSIS
            Search the users Desktop for another file called SRR.xlsx
        
        .DESCRIPTION
            Looks for a previous copy of the output file.  If found allows the user to either delete it, or rename it
            Renaming the output file will add a timestamp of the current time to the file name
        
        .INPUTS
            Must supply the $local_path -> get_local_path with $user will result in this required input
        
        .OUTPUTS
            Does not return anything, instead upon failure to delete previous file will exit the program instead
    #>
    $outputpath = $local_path+"\srr.xlsx"
    $test = Test-Path $outputpath
    if ($test){
        $status = [System.Windows.MessageBox]::Show("Previous SRR file found!`r`n`r`nWould you like to delete(Yes) or rename(No) this file.`r`n`r`nCancel will abort this program",'Error Message','YesNoCancel')
        if ($status -eq "Yes"){
            Remove-Item $outputpath -Force
        }
        elseif ($status -eq "No"){
            $add_date = Get-Date -Format MM-dd-yyy-HH-mm
            $filename = "srr" + $add_date + ".xlsx"
            Move-Item -Path $outputpath -Destination (Join-Path $local_path -childpath $filename)
        }
        else{
            Read-Host -Prompt "This program will now exit as you selected cancel and the program is designed not to overwrite previous files"
            exit
        }
    }
}

#Allows users to input sample name and/or SRR to search for
function get_user_input() {
    <#
        .SYNOPSIS
            Allows user to submit line by line either sample name or SRR #
        
        .DESCRIPTION
            User can submit line by line or copy and paste from a table
            Will filter out any repeats as with providing two warnings
                First warning is when the item is first submitted
                Second warning is for batch users will provide a single area of all the items found duplicated in this run
        
        .OUTPUTS
            Returns an ArrayList of all the items the user submitted
    #>
    $error_report = [System.Collections.ArrayList]@()
    $user_requests = [System.Collections.ArrayList]@()
    while($true){
        $user_input = Read-Host -Prompt "Enter data, must be one per line and blank line will start excel program"
        if ($user_input -ne '' -and  $user_requests -notcontains $user_input){
           $user_requests.add($user_input) | Out-Null
        }
        elseif ($user_input -ne '' -and  $user_requests -contains $user_input){
            Write-Host "Duplicate was found and removed"
            $error_report.add($user_input)  | Out-Null
        }
        else{
            break
        }
    }
    if ($error_report.Count -ne 0){
        Write-Host $error_report.Count
        Write-Host "One or more samples entered were found already present in this list and the duplicates have been removed."
        foreach ($item in $error_report){
            Write-Host $item "duplicate was found and removed in this process"
        }
    }
    return $user_requests
}


function get_website([String]$website_url){
    <#
        .SYNOPSIS
            Invoke-WebRequest of the required URL
        
        .INPUTS
            Must supply the $website_url that will be invoked
        
        .OUTPUTS
            Returns the the HMTL code as a String or Null if anything other than a status 200 is returned by the invoke method
    #>
    Write-Host "Looking for" $website_url
    $WebResponse = Invoke-WebRequest $website_url
    $html_code =  $WebResponse.Content
    $status =  [int]$WebResponse.StatusCode
    if ($status -eq 200){
        return $html_code
    }
    else {
        return $null
    }
}


function check_not_found($html){
    <#
        .SYNOPSIS
            Checks to see if more than 1 result is found 
        
        .DESCRIPTION
            At the time this was created, only 1 result was expected and no known automated method could be used to check a list of results, thus manual searching of this is required
            The data in the Excel sheet is still proceed with Not Found entered in all but the supplyed information and website.  The website link is provided for easy manual searching
        
        .INPUTS
            Must supply the $html -> get_website with $website_url will result in this required input
        
        .OUTPUTS
           If "Search results" is found in the HTML code, it assumed more than 1 result was found and thus returns true, otherwise false
    #>
    $split = $html -split "Search results"
    if ($split.length -eq 2){
        return $true
    }
    return $false
}

function get_SRR($html){
    <#
        .SYNOPSIS
            Uses the HMTL code to find the SRR number
            Known "bug" -> by design
                Only 1 SRR was expected at the time of creation and nothing has yet been programmed to handle many. 
                (Wish listed to check for more than 1)
        
        .INPUTS
            Must supply the $html -> get_website with $website_url will result in this required input
        
        .OUTPUTS
            returns SRR # as String
    #>
    $split = $html -split "SRR" #split the data stream
    $srr = $split[1]              #select 2nd part
    $srr = $srr -split "",10       #split into 9 parts 
    $srr = $srr[0..8]             #keep first 8
    $output = "SRR"
    foreach ($item in $srr){      #checks that a shorter SRR (legacy) isn't provided
        if ($item -ne '"'){
            $output += $item
        }
        else{
            break
        }
    }
    return $output
}

function get_SAMN_sample_ID($html){
    <#
        .SYNOPSIS
            Uses the HMTL code to find the Sample Name
        
        .INPUTS
            Must supply the $html -> get_website with $website_url will result in this required input
        
        .OUTPUTS
            returns Sample Name as String
    #>
    $samn_finder = $html -split "Sample: <span>"
    $samn_finder = $samn_finder[1] -split "SAMN"
    $samn_finder = $samn_finder[1] -split '"'
    $samn_finder = "SAMN"+$samn_finder[0]
    $website_url = "https://www.ncbi.nlm.nih.gov/biosample/" + $samn_finder
    $html = get_website($website_url)
    $sample_finder = $html -split "Sample Name:"
    $sample_finder = $sample_finder[1] -split "SRA"
    $sample_finder = $sample_finder[0]
    $sample_finder = $sample_finder -replace ";",""
    return $sample_finder
}

function get_bio_project($html){
    <#
        .SYNOPSIS
            Uses the HMTL code to find the bio_project
            Possible "bug"
                Bio Projects start with PRJNA for our case use and thus was programmed as such
        
        .INPUTS
            Must supply the $html -> get_website with $website_url will result in this required input
        
        .OUTPUTS
            returns bio_project as String
    #>
    $bioproject = $html -split "PRJNA"
    $bioproject = $bioproject[1]
    $bioproject = $bioproject -split ""
    $bioproject = $bioproject[0..6]
    $bioproject_complete = "PRJNA"
    foreach ($item in $bioproject){
        if ($item -eq '"'){
            break
        }
        else{
            $bioproject_complete += $item
        }
    }
    return $bioproject_complete
}

function get_submitted_by($html){
    <#
        .SYNOPSIS
            Uses the HMTL code to find the submitter of the sample
        
        .INPUTS
            Must supply the $html -> get_website with $website_url will result in this required input
        
        .OUTPUTS
            returns the submitter as String
    #>
    $submittedby = $html -split "Submitted by:"
    $submittedby = $submittedby[1] -split "Study:"
    $submittedby = $submittedby[0] -split "<span>"
    $submittedby = $submittedby[1] -split "<"
    $submittedby = $submittedby[0]
    return $submittedby
}

function program(){
    <#
        .SYNOPSIS
           The main program
        
        .DESCRIPTION
            This is the meat of the program.  It calls all the required methods for normal operations and assembles the Excel sheet at the end
        
        .OUTPUTS
            Excel spreadsheet with results
    #>

    #define starting items
    [System.Collections.ArrayList]$samples = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$srrs = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$bioprojects = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$submitters = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$websites = [System.Collections.ArrayList]@()

    #build basic program requirements
    $user = get_user
    $local_path = get_local_path $user
    $outputpath = $local_path+"\srr.xlsx"
    search_prev_file $local_path

    #get user input
    $user_input = get_user_input

    #check to see if the user suppled at least 1 item
    if ($user_input.Count -gt 0){

        #loop through each item the user supplied to gather results
        foreach ($item in $user_input){
            $srr_provided = $false
            $sample_provided = $false

            #sample or SRR provided by user?
            $item_check = $item -split "SRR"
            if($item_check.length -gt 1){
                $srr_provided = $true
            }
            else{
                $sample_provided = $true
            }

            #build path and attempt to connect and get HTML code
            $website_path = "https://www.ncbi.nlm.nih.gov/sra/?term="
            $website_url = $website_path + $item
            $html = get_website $website_url

            #if html returns false aka no code, then fill default info
            $status = check_not_found $html
            if ($html -eq $false -or $status -eq $true ){
                if ($srr_provided){
                    $srrs.add($item)  | Out-Null
                }
                else{
                    $srrs.add("Not Found")  | Out-Null
                }
                if ($sample_provided){
                    $samples.add($item) | Out-Null
                }
                else{
                    $srrs.add("Not Found") | Out-Null
                }
                $bioprojects.add("Not Found") | Out-Null
                $submitters.add("Not Found") | Out-Null
                $websites.add($website_url) | Out-Null
            }

            #$html returned with code to process
            else{

                #if SRR was provided by user then add by default SRR and search for sample name
                if ($srr_provided){
                    $sample = get_SAMN_sample_ID $html
                    $samples.add($sample) | Out-Null
                    $srrs.add($item) | Out-Null
                }

                #if sample was provided by user then add by default sample and search for SRR

                if ($sample_provided){
                    $srr = get_SRR $html
                    $srrs.add($srr) | Out-Null
                    $samples.add($item) | Out-Null
                }

                #find bio_projects, submittters and add them
                $bioproject = get_bio_project $html
                $bioprojects.add($bioproject) | Out-Null
                $submittedby = get_submitted_by $html
                $submitters.add($submittedby) | Out-Null

                #add website
                $websites.add($website_url) | Out-Null
            }    
        }
    

        #prepare to write data
        if ($samples.Count -gt 0 -and ($samples.Count -eq $bioprojects.Count -and $samples.Count -eq $submitters.Count -and $samples.Count -eq $websites.Count -and $samples.Count -eq $srrs.Count )){
            Write-Host "Writing srr.xlsx..."
            $excel = New-Object -ComObject excel.application
            $excel.visible = $false
            $workbook = $excel.Workbooks.Add()
            $uregwksht= $workbook.Worksheets.Item(1)
            $uregwksht.Name = 'Test'
            $row = 1
            $column = 1
            #create titles
            $titles = 'Sample_name', 'SRR #', 'Bio Project', 'Submitted By', 'URL'
            foreach ($item in $titles){
                $uregwksht.Cells.Item($row,$column)= $item
                $column+=1
            }
            $row+=1
            $column = 1
            $local_itr = 0
            while ($local_itr -lt $srrs.Count){
                $uregwksht.Cells.Item($row,$column) = " " + $samples[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $srrs[$local_itr]
                if ($srr[$local_itr] -eq "not found"){
                    Write-Host $testers[$local_itr] "was not found"
                }
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $bioprojects[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $submitters[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $websites[$local_itr]
                $row+=1
                $column = 1
                $local_itr+=1
            }
            $workbook.SaveAs($outputpath)
            $excel.Quit()
            Read-Host -Prompt "Program has completed, you may now exit this window by pressing enter closing the screen."
        }
        else {
            Read-Host -Prompt "A critcal errror occured and to protect the intergity of the report, noting was generated.  The program failed to build an even amount of data points which would result in row mismatch"
        }
        
    }

    #user failed to supply at least 1 input and will notify the user and start the program again
    else{
        Write-Host "You must supply at least 1 item for the program to work."
        program
    }

}
program
