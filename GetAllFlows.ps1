param(
    [Parameter(mandatory=$true)]
    [string] $LogPath
)

try{
    #Verify if MSOnline Module is available
    if (Get-Module -ListAvailable -Name MSonline) {
        #Import MSOnline Module
        import-module MSOnline -ErrorAction SilentlyContinue

        #Verify if the Microsoft.PowerApps.Administration.PowerShell Module and Microsoft.PowerAPps.PowerShell are installed
        if (Get-Module -ListAvailable -Name Microsoft.PowerApps.Administration.PowerShell) {
            if (Get-Module -ListAvailable -Name Microsoft.PowerApps.PowerShell) {
        
                #Test if logpath exists
                If(Test-Path $LogPath) { 
                    #Start script
                    Try{
                        #Object collections
                        $flowCollection = @()
                        $environmentCollection = @()

                        #Connect to the correct O365 Tenant
                        Connect-MsolService 
                        
                        #Connect to the Flow Environment
                        Add-PowerAppsAccount

                        #Retrieve all users UPN and ID
                        $users = Get-MsolUser -All | Select-Object UserPrincipalName, ObjectId | Sort-Object DisplayName

                        #Retrieve all flow environments
                        $environments = Get-FlowEnvironment | Sort-Object EnvironmentName
                        
                        #Retrieve all flows
                        $flows = get-AdminFlow | Sort-Object EnvironmentName
                      
                        #loop through all environments
                        foreach($environment in $environments){
                            #fill the collection with information
                            $envProperties = $environment.internal.properties
                            [datetime]$createdTime = $envProperties.createdTime
                            $environmentCollection += new-object psobject -property @{displayName = $envProperties.displayName;InternalName = $environment.EnvironmentName;SKU = $envProperties.environmentSku;EnvType = $envProperties.environmentType;Region = $envProperties.azureRegionHint;Created = $createdTime;CreatedBy = $envProperties.createdby.displayname}
                        }

                        #loop through all flows
                        foreach($flow in $flows){
                            #fill the collection with information
                            $flowProperties = $flow.internal.properties
                            $creatorName = $users | where-object{$_.ObjectId -eq $flowProperties.creator.UserID}
                            
                            $triggers = $flowProperties.definitionsummary.triggers
                            $actions = $flowProperties.definitionsummary.actions | where-object {$_.swaggerOperationId}
                            
                            $triggerString = ""
                            foreach($trigger in $triggers){
                                if($triggerString -ne ""){
                                    $triggerString = $triggerString + "<br />"
                                }
                                $triggerString = $triggerString + "$($trigger.swaggerOperationId)"
                            }
                            
                            $actionsString = ""
                            foreach($action in $actions){
                                if($actionsString -ne ""){
                                    $actionsString = $actionsString + "<br />"
                                }
                                $actionsString = $actionsString + "$($action.swaggerOperationId)"
                            }
                            
                            
                            [nullable[datetime]]$modifiedTime = $flow.LastModifiedTime
                            [nullable[datetime]]$createdTime = $flowProperties.createdTime
                            
                            $flowCollection += new-object psobject -property @{displayName = $flowProperties.displayName;environment = $flowProperties.Environment.name;State = $flowProperties.State;Triggers = $triggerString;Actions = $actionsString;Created = $createdTime;Modified = $modifiedTime;CreatedBy = $creatorName.userPrincipalName}
                        }    

                        #We now have our collections so we are building the HTML page to get a direct view
                        #List of all Flow environments
                        $article = "<h2>List of all Flow environments</h2>"
                        $article += "<table>
                                    <tr>
                                        <th>displayName</th>
                                        <th>InternalName</th>
                                        <th>SKU</th>
                                        <th>Type</th>
                                        <th>Region</th>
                                        <th>Created</th>
                                        <th>CreatedBy</th>
                                    </tr>"
                        
                        foreach($environmentColl in $environmentCollection){
                        $article += "<tr>
                                        <td>$($environmentColl.displayName)</td>
                                        <td>$($environmentColl.InternalName)</td>
                                        <td>$($environmentColl.SKU)</td>
                                        <td>$($environmentColl.EnvType)</td>
                                        <td>$($environmentColl.Region)</td>
                                        <td>$($environmentColl.Created)</td>
                                        <td>$($environmentColl.CreatedBy)</td>
                                    </tr>"
                        }
                        
                        $article += "</table>"

                        #List of all Flows
                        $article += "<h2>List of all Flows</h2>"
                        $article += "<table>
                                    <tr>
                                        <th>displayName</th>
                                        <th>environment</th>
                                        <th>State</th>
                                        <th>Triggers</th>
                                        <th>Actions</th>
                                        <th>Created</th>
                                        <th>Modified</th>
                                        <th>CreatedBy</th>
                                    </tr>"
                        
                        foreach($flowColl in $flowCollection){
                        $article += "<tr>
                                        <td>$($flowColl.displayName)</td>
                                        <td>$($flowColl.environment)</td>
                                        <td>$($flowColl.State)</td>
                                        <td>$($flowColl.Triggers)</td>
                                        <td>$($flowColl.Actions)</td>
                                        <td>$($flowColl.Created)</td>
                                        <td>$($flowColl.Modified)</td>
                                        <td>$($flowColl.CreatedBy)</td>
                                    </tr>"
                        }
                        
                        $article += "</table>"

                        $date = get-date
                        $today = $date.ToString("ddMMyyyy_HHmm")
                        $LogPath = Join-Path $LogPath "HTMLFlowReport_$($today).html"    
                        
                        #Head
                        $head = "
                        <html xmlns=`"http://www.w3.org/1999/xhtml`">
                            <head>
                                <style>
                                    @charset `"UTF-8`";
 
                                    @media print {
                                        body {-webkit-print-color-adjust: exact;}
                                    }
                         
                                    div.container {
                                        width: 100%;
                                        border: 1px solid gray;
                                    }
                                     
                                    header {
                                        padding: 0.1em;
                                        color: white;
                                        background-color: #000033;
                                        color: white;
                                        clear: left;
                                        text-align: center;
                                        border-bottom: 2px solid #FF0066
                                    }
                                     
                                    footer {
                                        padding: 0.1em;
                                        color: white;
                                        background-color: #000033;
                                        color: white;
                                        clear: left;
                                        text-align: center;
                                        border-top: 2px solid #FF0066
                                    }
 
                                    article {
                                        margin-left: 20px;
                                        min-width:600px;
                                        min-height: 600px;
                                        padding: 1em;
                                    }
                                     
                                    th{
                                        border:1px Solid Black;
                                        border-Collapse:collapse;
                                        background-color:#000033;
                                        color:white;
                                    }
                                     
                                    th{
                                        border:1px Solid Black;
                                        border-Collapse:collapse;
                                    }
                                     
                                    tr:nth-child(even) {
                                      background-color: #dddddd;
                                    }
 
                                </style>
                            </head>
                        "
                        
                        #Header
                        $date = (get-date).tostring("dd-MM-yyyy")
                        $header = "
                            <h1>Flow Report</h1>
                            <h5>$($date)</h5>
                        "
                        
                     #   #Footer
                        $Footer = "
                            Copyright &copy;
                        "
                        
                        #Full HTML
                        $HTML = "
                            $($Head)
                            <body class=`"Inventory`">
                                <div class=`"container`">
                                    <header>
                                        $($Header)
                                    </header>
                                     
                                    <article>
                                        $($article)
                                    </article>
                                             
                                    <footer>
                                        $($footer)
                                    </footer>
                                </div>
                            </body>
                            </html>
                        " 
                        add-content $HTML -path $LogPath

                        Write-Host "Flow overview created at $($LogPath), it will also open automatically in 5 seconds" -foregroundcolor green

                        start-sleep -s 5
                        Invoke-Item $LogPath
                    }
                    catch{
                        write-host "Error occurred: $($_.Exception.Message), please post this error on https://www.cloudsecuritea.com" -foregroundcolor red
                    }
                } Else { 
                    Write-Host "The path $($LogPath) could not be found. Please enter a correct path to store the Office 365 subscription and license overview" -foregroundcolor yellow
                }
            }
            Else{Write-Host "The new Microsoft.PowerApps.PowerShell Module is not installed. Please install using the link in the blog" -foregroundcolor yellow}
        }
        Else{Write-Host "The new Microsoft.PowerApps.Administration.PowerShell Module is not installed. Please install using the link in the blog" -foregroundcolor yellow}
    } else {
        Write-Host "MSOnline module not loaded. Please install the MSOnline module with Install-Module MSOnline" -foregroundcolor yellow
    }
}
catch{
    write-host "Error occurred: $($_.Exception.Message)" -foregroundcolor red
}