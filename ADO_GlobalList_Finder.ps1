# [CmdletBinding(DefaultParameterSetName = 'InstanceURI')]
# param (
# [Parameter(Mandatory, ParameterSetName = 'InstanceURI', HelpMessage = "Enter Instance URL")]
# [ValidateNotNullOrEmpty()]
# [string] $tfsInstanceURI
# )

param (
[Parameter(Mandatory, ParameterSetName = 'InstanceURI')]
[ValidateNotNullOrEmpty()]
[string] $tfsInstanceURI = $(Read-Host "Enter Instance URL")
)

#Add TFS assemblies
Add-Type -Path "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Microsoft.TeamFoundation.Client.dll"
Add-Type -Path "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Microsoft.TeamFoundation.WorkItemTracking.Client.dll"

# Define variables
$filesStoreLocation = (Get-Item -Path ".\").FullName #Current folder path to store files
$tfsConfigurationServer = [Microsoft.TeamFoundation.Client.TfsConfigurationServerFactory]::GetConfigurationServer($tfsInstanceURI)
$tpcService = $tfsConfigurationServer.GetService("Microsoft.TeamFoundation.Framework.Client.ITeamProjectCollectionService")
$witadminLocation = "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE"

$sortedCollections = $tpcService.GetCollections() | Sort-Object -Property Name
$numberOfProjects = 0

foreach($collection in $sortedCollections) {
    $teamProjectCollectionName = $collection.Name
    $collectionUri = $tfsInstanceURI + "/" + $teamProjectCollectionName
    $tfsTeamProject = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($collectionUri)
    $cssService = $tfsTeamProject.GetService("Microsoft.TeamFoundation.Server.ICommonStructureService3")   
    $sortedProjects = $cssService.ListProjects() | Sort-Object -Property Name

    foreach($project in $sortedProjects)
    {
        $numberOfProjects++

        $tfs = [Microsoft.TeamFoundation.Client.TeamFoundationServerFactory]::GetServer($tfsInstanceURI)

        $type = [Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore]

        $store = [Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore] $tfs.GetService($type)


        $workItem = new-object Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem($store.Projects[0].WorkItemTypes[0])


        $workItemTypes = $store.Projects[$project.Name].WorkItemTypes

        Write-Host "Total WI Types:" $workItemTypes.Count
        $teamProjectName = $store.Projects[$project.Name].Name
        
        for ($int = 0; $int -lt $workItemTypes.Count; $int++)
        {
            $WITName = $store.Projects[$project.Name].WorkItemTypes[$int].Name
            
            CD $witadminLocation
            .\witadmin.exe exportwitd /collection:`"$collectionUri`" /p:`"$teamProjectName`" /n:`"$WITName`" /f:$filesStoreLocation\$teamProjectCollectionName_$teamProjectName-$WITName.xml
            Write-Host `"$collectionUri`" /p:`"$teamProjectName`" /n:`"$WITName`" /f:$filesStoreLocation\$teamProjectCollectionName_$teamProjectName-$WITName.xml
        }
        
        .\witadmin.exe exportgloballist /collection:`"$collectionUri`" /f:$tempStore\GL_$teamProjectCollectionName.xml

        $collectionGlobalLists = [xml](Get-Content $tempStore\GL_$teamProjectCollectionName.xml)

        #Read all work item defination files

        Get-ChildItem * -Recurse -Filter "*.xml" |
        ForEach-Object {
        try
        {
            $content = [xml](Get-Content $_.FullName)
        }
        catch {}

        $content.selectNodes('//GLOBALLIST') | select name

        } | Export-Csv $tempStore\GL_Found_$teamProjectCollectionName.csv
        

    }
