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
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791703
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300061
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791715
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300073
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791735
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300093
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791736
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300094
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791737
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300095
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791752
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300110
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791755
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300113
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791758
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300116
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-0011
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-0357
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300125
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-0493
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300128
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-0585
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300131
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1331
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300135
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1368
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300136
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1405
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300139
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1463
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300140
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1531
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300143
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Looking for https://www.ncbi.nlm.nih.gov/sra/?term=02-1579
            Looking for https://www.ncbi.nlm.nih.gov/biosample/SAMN03300146
            Looking for https://www.ncbi.nlm.nih.gov/bioproject/251692
            Writing srr.xlsx...
            Program has completed, you may now exit this window by pressing enter closing the screen.:

        Output file reuslts ->
            Website	                                            Sample Name	SRR #	    Organism	                                Bioproject #	Bioproject Name	                                                    Submitter	                                                                                                                            Publication	            Publication Link	Attributes	    Strain	Isolate	        Host	    Isolate Source	Collection Date	Geographic Location	Sample Type
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791702	  01-0467 	SRR1791702	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-0467 	 MB11 	    Cattle 	    Tissue 	        2001	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791703	  01-0843 	SRR1791703	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-0843 	 MB12 	    Cattle 	    Tissue 	        2001	        Mexico          	Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791715	  01-2374 	SRR1791715	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-2374 	 MB24 	    Cattle 	    Tissue 	        2001	        USA:MO	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791735	  01-4106 	SRR1791735	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-4106 	 MB44 	    Cattle 	    Tissue 	        2001	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791736	  01-4280 	SRR1791736	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-4280 	 MB45 	    Cattle 	    Tissue      	2001	        USA:CO	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791737	  01-4283 	SRR1791737	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-4283 	 MB46 	    Cattle 	    Tissue 	        2001	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791752	  01-5745 	SRR1791752	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-5745 	 MB61 	    Cattle 	    Tissue 	        2001	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791755	  01-6106 	SRR1791755	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-6106 	 MB64 	    Cattle 	    Tissue 	        2001	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=SRR1791758	  01-6318 	SRR1791758	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            01-6318 	 MB67 	    Cattle 	    Tissue 	        2001	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-0011	      02-0011	Not Found	Not Found	                                Not Found	    Not Found	                                                        Not Found	                                                                                                                            Not Found	            Not Found		                    Not Found	 Not Found	Not Found	 Not Found	    Not Found	    Not Found	        Not Found
            https://www.ncbi.nlm.nih.gov/sra/?term=02-0357	      02-0357	SRR1791767	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-0357 	 MB76 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-0493	      02-0493	SRR1791770	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-0493 	 MB79 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-0585	      02-0585	SRR1791773	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-0585 	 MB82 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-1331	      02-1331	SRR1791777	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-1331 	 MB86 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-1368	      02-1368   SRR1791778	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-1368 	 MB87 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-1405	      02-1405   SRR1791781	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-1405 	 MB90 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-1463	      02-1463	SRR1791782	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-1463 	 MB91 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-1531	      02-1531	SRR1791785	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-1531 	 MB94 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided
            https://www.ncbi.nlm.nih.gov/sra/?term=02-1579	      02-1579	SRR1791788	Mycobacterium tuberculosis variant bovis	PRJNA251692	    United States Department of Agriculture Mycobacterium Diagnostics	USDA Animal Plant Health Inspection Service-National Veterinary Services Laboratory-Diagnostic Bacteriology Laboratory (USDA-NVSL-DBL)	No Publication Found	No Link Provided		            02-1579 	 MB97 	    Cattle 	    Tissue 	        2002	        Mexico	            Not Provided

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

function get_organsim($html){
    <#
        .SYNOPSIS
            Gets the organism
        
        .INPUTS
            Must supply the $html
        
        .OUTPUTS
            Returns the organism as String
    #>
    $organism = $html -split "Organism:"
    $organism = $organism[1] -split 'href="'
    $organism = $organism[1] -split '">',2
    $organism = $organism[1] -split "</a>"
    $organism = $organism[0]
    return $organism
}

function get_attribute_code($SAMN_html){
    <#
        .SYNOPSIS
            filters down the code (html) to just the required section to speed up the overall process
        
        .INPUTS
            Must supply the $SAMN_html
        
        .OUTPUTS
            Returns the section code (html) as String
    #>
    $code = $SAMN_html -split "<dt>Attributes</dt>"
    $code = $code[1] -split "<dt>BioProject</dt>"
    $code = $code[0]
    return $code
}

function get_attribute_isolate($attr_code){
    <#
        .SYNOPSIS
            Gets the isolate if provided
        
        .INPUTS
            Must supply the $attr_code
        
        .OUTPUTS
            Returns the isolate as String
    #>
    $isolate = $attr_code -split "<th>isolate</th><td>"
    if ($isolate.length -eq 2){
        $isolate = $isolate[1] -split '<',2
        $isolate = $isolate[0]
        return $isolate
    }
    return "Not Provided"
}

function get_attribute_strain($attr_code){
    <#
        .SYNOPSIS
            Gets the strain if provided
        
        .INPUTS
            Must supply the $attr_code
        
        .OUTPUTS
            Returns the string as String
    #>
    $strain = $attr_code -split "<th>strain</th><td>"
    if ($strain.length -eq 2){
        $strain = $strain[1] -split '<',2
        $strain = $strain[0]
        return $strain
    }
    return "Not Provided"
}

function get_attribute_host($attr_code){
    <#
        .SYNOPSIS
            Gets the host if provided
        
        .INPUTS
            Must supply the $attr_code
        
        .OUTPUTS
            Returns the host as String
    #>
    $host_name = $attr_code -split "<th>host</th><td>"
    if ($host_name.length -eq 2){
        $host_name = $host_name[1] -split '<',2
        $host_name = $host_name[0]
        return $host_name
    }
    return "Not Provided"
}

function get_attribute_geographic_location($attr_code){
    <#
        .SYNOPSIS
            Gets the geographic location if provided
        
        .INPUTS
            Must supply the $attr_code
        
        .OUTPUTS
            Returns the geographic location as String
    #>
    $geographic_location = $attr_code -split "<th>geographic location</th><td>"
    if ($geographic_location.length -eq 2){
        $geographic_location = $geographic_location[1] -split ' ref="'
        $geographic_location = $geographic_location[1] -split '>',2
        $geographic_location = $geographic_location[1] -split '<',2
        $geographic_location = $geographic_location[0]
        return $geographic_location
    }
    return "Not Provided"
}

function get_attribute_sample_type($attr_code){
    <#
        .SYNOPSIS
            Gets the sample type if provided
        
        .INPUTS
            Must supply the $attr_code
        
        .OUTPUTS
            Returns the sample type as String
    #>
    $sample_type = $attr_code -split "<th>sample type</th><td>"
    if ($sample_type.length -eq 2){
        $sample_type = $sample_type[1] -split '<',2
        $sample_type = $sample_type[0]
        return $sample_type
    }
    return "Not Provided"
}

function get_attribute_isolation_source($attr_code){
    <#
        .SYNOPSIS
            Gets the isolation source if provided
        
        .INPUTS
            Must supply the $attr_code
        
        .OUTPUTS
            Returns the isolation_source as String
    #>
    $isolation_source = $attr_code -split "<th>isolation source</th><td>"
    if ($isolation_source.length -eq 2){
        $isolation_source = $isolation_source[1] -split '<',2
        $isolation_source = $isolation_source[0]
        return $isolation_source
    }
    return "Not Provided"
}

function get_attribute_collection_date($attr_code){
    <#
        .SYNOPSIS
            Gets the collection date if provided
        
        .INPUTS
            Must supply the $attr_code
        
        .OUTPUTS
            Returns the collection_date as String
    #>
    $collection_date = $attr_code -split "<th>collection date</th><td>"
    if ($collection_date.length -eq 2){
        $collection_date = $collection_date[1] -split '<',2
        $collection_date = $collection_date[0]
        return $collection_date
    }
    return "Not Provided"
}

function get_samn_website($html){
    <#
        .SYNOPSIS
            Gets the samn webstie code (html)
        
        .INPUTS
            Must supply the $html 
        
        .OUTPUTS
            Returns the samn website code (html) as String
    #>
    $samn_finder = $html -split "Sample: <span>"
    $samn_finder = $samn_finder[1] -split "SAMN"
    $samn_finder = $samn_finder[1] -split '"'
    $samn_finder = "SAMN"+$samn_finder[0]
    $website_url = "https://www.ncbi.nlm.nih.gov/biosample/" + $samn_finder
    $SAMN_html = get_website($website_url)
    return $SAMN_html
}

function get_sample_name($SAMN_html){
    <#
        .SYNOPSIS
            Uses the HMTL code to find the Sample Name
        
        .INPUTS
            Must supply the $html -> get_website with $website_url will result in this required input
        
        .OUTPUTS
            returns Sample Name as String
    #>
    $sample_finder = $SAMN_html -split "Sample Name:"
    $sample_finder = $sample_finder[1] -split "SRA"
    $sample_finder = $sample_finder[0]
    $sample_finder = $sample_finder -replace ";",""
    return $sample_finder
}

function get_bioproject_html($bioproject_number){
    <#
        .SYNOPSIS
            gets the bioproject code (html)
        
        .INPUTS
            Must supply the $bioproject_number
        
        .OUTPUTS
            Returns the bioproject code (html) as String
    #>
    $bio_num = $bioproject_number -replace "PRJNA",""
    $website_url = "https://www.ncbi.nlm.nih.gov/bioproject/" + $bio_num
    $bio_html = get_website $website_url
    return $bio_html

}

function get_bio_project_number($SAMN_html){
    <#
        .SYNOPSIS
            Uses the HMTL code to find the bio_project_number
            Possible "bug"
                Bio Projects start with PRJNA for our case use and thus was programmed as such
        
        .INPUTS
            Must supply the $html -> get_website with $website_url will result in this required input
        
        .OUTPUTS
            returns bio_project_number as String
    #>
    $bioproject = $SAMN_html -split "PRJNA"
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

function get_bioproject_code($bioproject_html){
    <#
        .SYNOPSIS
            Gets the section of bioproject code (html) for faster processing
        
        .INPUTS
            Must supply the $bioproject_html
        
        .OUTPUTS
            Returns the section of code (html) as String
    #>
    $code = $bioproject_html -split '<div id="maincontent"'
    return $code[1]
}

function get_bio_project_name($bioproject_code){
    <#
        .SYNOPSIS
            Gets the bioproject title (name)
        
        .INPUTS
            Must supply the $bioproject_code
        
        .OUTPUTS
            Returns the bioproject name as String
    #>
    $title = $bioproject_code -split '<div class="Description">'
    $title = $title[1] -split "<"
    return $title[0]
  
}

function get_publication_link($bioproject_code) {
    <#
        .SYNOPSIS
            Gets the publication if provided
        
        .INPUTS
            Must supply the $bioproject_code
        
        .OUTPUTS
            Returns the link as a string
    #>
    $link = $bioproject_code -split "Publications"
    if ($link.length -eq 2){
        $link = $link[1] -split 'href="',2
        $link = $link[1] -split '"'
        $link = $link[0]
        $link = "https://www.ncbi.nlm.nih.gov" + $link
        return $link
    }
    return "No Link Provided"
}

function get_publication_info($bioproject_code){
    <#
        .SYNOPSIS
            Gets the publication title if provided
        
        .INPUTS
            Must supply the $bioproject_code
        
        .OUTPUTS
            Returns the title as a string
    #>
    $title = $bioproject_code -split "Publications"
    if ($title.length -eq 2){
        $title = $title[1] -split "</a>"
        $title = $title[1] -split "</td>"
        $title = $title[0] -replace '"',"" -replace "/","" -replace "<i>",""
        return $title
    }
    return "No Publication Found"
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
    [System.Collections.ArrayList]$websites = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$sample_names = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$srrs = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$organisms = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$bioproject_numbers = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$bioproject_names = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$submitters = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$publications = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$publication_links = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$strains = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$isolates = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$hosts = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$isolate_sources = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$collection_dates = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$geo_locs = [System.Collections.ArrayList]@()
    [System.Collections.ArrayList]$sample_types = [System.Collections.ArrayList]@()
    

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
                    $sample_names.add($item) | Out-Null
                }
                else{
                    $srrs.add("Not Found") | Out-Null
                }
                $bioproject_numbers.add("Not Found") | Out-Null
                $submitters.add("Not Found") | Out-Null
                $websites.add($website_url) | Out-Null
                $organisms.add("Not Found") | Out-Null
                $bioproject_names.add("Not Found") | Out-Null  
                $publications.add("Not Found") | Out-Null
                $publication_links.add("Not Found") | Out-Null  
                $strains.add("Not Found") | Out-Null  
                $isolates.add("Not Found") | Out-Null 
                $hosts.add("Not Found") | Out-Null 
                $isolate_sources.add("Not Found") | Out-Null 
                $collection_dates.add("Not Found") | Out-Null 
                $geo_locs.add("Not Found") | Out-Null 
                $sample_types.add("Not Found") | Out-Null
            }

            #$html returned with code to process
            else{
                #get samn website code
                $SAMN_html = get_samn_website $html

                #get bioproject hmtl and code and required bio_project_number
                $bioproject_number = get_bio_project_number $SAMN_html
                $bioproject_html = get_bioproject_html $bioproject_number
                $bioproject_code = get_bioproject_code $bioproject_html

                #if SRR was provided by user then add by default SRR and search for sample name
                if ($srr_provided){
                    $sample = get_sample_name $SAMN_html
                    $sample_names.add($sample) | Out-Null
                    $srrs.add($item) | Out-Null
                }

                #if sample was provided by user then add by default sample and search for SRR

                if ($sample_provided){
                    $srr = get_SRR $html
                    $srrs.add($srr) | Out-Null
                    $sample_names.add($item) | Out-Null
                }

                #find bio_project_numbers and names, submittters along with publicication and add them
                $bioproject_numbers.add($bioproject_number) | Out-Null
                $bioproject_name = get_bio_project_name $bioproject_code
                $bioproject_names.add($bioproject_name) | Out-Null
                $submittedby = get_submitted_by $html
                $submitters.add($submittedby) | Out-Null
                $publication = get_publication_info $bioproject_code
                $publications.add($publication) | Out-Null
                $publication_link = get_publication_link $bioproject_code
                $publication_links.add($publication_link) | Out-Null

                #add website
                $websites.add($website_url) | Out-Null

                #add organismn
                $ogranism = get_organsim $html
                $organisms.add($ogranism) | Out-Null

                #add attributes code and the attributes themselves
                $attr_code = get_attribute_code $SAMN_html
                $strain = get_attribute_strain $attr_code
                $strains.add($strain) | Out-Null 
                $isolate = get_attribute_isolate $attr_code
                $isolates.add($isolate) | Out-Null 
                $host_name = get_attribute_host $attr_code
                $hosts.add($host_name) | Out-Null
                $isolate_source = get_attribute_isolation_source $attr_code
                $isolate_sources.add($isolate_source) | Out-Null
                $collection_date = get_attribute_collection_date $attr_code
                $collection_dates.add($collection_date) | Out-Null
                $geo_loc = get_attribute_geographic_location $attr_code
                $geo_locs.add($geo_loc) | Out-Null 
                $sample_type = get_attribute_sample_type $attr_code
                $sample_types.add($sample_type) | Out-Null 
            }    
        }
    

        #prepare to write data
        if ($sample_names.Count -gt 0 -and ($sample_names.Count -eq $bioproject_numbers.Count -and $sample_names.Count -eq $submitters.Count -and $sample_names.Count -eq $websites.Count -and $sample_names.Count -eq $srrs.Count )){
            Write-Host "Writing srr.xlsx..."
            $excel = New-Object -ComObject excel.application
            $excel.visible = $false
            $workbook = $excel.Workbooks.Add()
            $uregwksht= $workbook.Worksheets.Item(1)
            $uregwksht.Name = 'Test'
            $row = 1
            $column = 1
            #create titles
            $titles = 'Website,Sample Name,SRR #,Organism,Bioproject #,Bioproject Name,Submitter,Publication,Publication Link,Attributes,Strain,Isolate,Host,Isolate Source,Collection Date,Geographic Location,Sample Type'
            $titles = $titles -split ","
            foreach ($item in $titles){
                $uregwksht.Cells.Item($row,$column)= $item
                $column+=1
            }
            $row+=1
            $column = 1
            $local_itr = 0
            while ($local_itr -lt $srrs.Count){
                $uregwksht.Cells.Item($row,$column) = $websites[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = " " + $sample_names[$local_itr]
                #if ($srr[$local_itr] -eq "not found"){
                #    Write-Host $testers[$local_itr] "was not found"
                #}
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $srrs[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $organisms[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $bioproject_numbers[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $bioproject_names[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $submitters[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $publications[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $publication_links[$local_itr]
                $column+=1
                #attributes left blank by design
                $column+=1
                $uregwksht.Cells.Item($row,$column) = " " + $strains[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = " " + $isolates[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $hosts[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = " " + $isolate_sources[$local_itr]
                $column+=1  
                $uregwksht.Cells.Item($row,$column) = $collection_dates[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $geo_locs[$local_itr]
                $column+=1
                $uregwksht.Cells.Item($row,$column) = $sample_types[$local_itr]
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
