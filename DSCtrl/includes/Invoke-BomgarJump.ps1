function Invoke-BomgarJump {
    <#
        .SYNOPSIS
        Provides a UI for selecting computer objects from a list
        .DESCRIPTION
        When this function is provided a collection of ADComputer objects, it displays that collection in a Windows form; allowing the end user to select from the list using checkboxes.

        When the Accept/OK button is clicked, the function filters the list based on which checkboxes were selected and returs the filtered list

        NOTE: While the default behavior is to return the list of SELECTED objects, you can specify the -Invert parameter to return a list of unselected objects

        In the event that the form is closed by any other means than the Accept/OK button, this function will throw an error and halt execuion
        .EXAMPLE
        Get-ADComputer -SearchBase (Get-ADRootDSE).defaultNamingContext -Filter "name -like 'UTO*'" | Select-ADComputer
        .EXAMPLE
        Select-ADComputer -ADComputers $labComputers -Invert | Format-List name
        .PARAMETER ADComputers
        A collection (array) of ADComputer objects. This collection will be presented to the user, who will have an opportunity to select which members they would like to filter.
        .PARAMETER invertSelection
        The default behavior of this function is to return the items the user selected from the collection. Specifying this switch changes that; and will result in the function returning the unselected items.
        .NOTES
        When calling this function, be aware that it will throw an exception and halt execution if the user closes the window or uses the cancel button; resulting in Null being returned in the data pipeline.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Computer name to jump to')]
        [Alias('Computer')]
        [String]$ComputerName
    )

    # find Bomgar installation
    $consolepath = $null

    # Check Program Files First
    $bomgardir = Join-Path $ENV:ProgramFiles 'bomgar'
    if (Test-Path -Path $bomgarDir) {
        $consolepath = Get-ChildItem -Path $bomgarDir -Recurse | Where-Object {$_.name -ilike 'bomgar-rep*.exe'} | Select-Object -Last 1
    }

    # Also check x86 if that failed
    $bomgardir = Join-Path ${ENV:ProgramFiles(x86)} 'bomgar'
    if ( ($consolepath -eq $null) -and (Test-Path -Path $bomgarDir) ) {
        $consolepath = Get-ChildItem -Path $bomgarDir -Recurse | Where-Object {$_.name -ilike 'bomgar-rep*.exe'} | Select-Object -Last 1
    }

    # If we still could not find the console, then return an error code and let the caller figure it out
    if ($consolepath -eq $null) {
        return 'Bomgar client not found'
    }

    # Launch the client with the script arguments
    &"$($consolepath[0].fullname)" '--run-script' "`"action=start_pinned_client_session&search_string=${ComputerName}`""
    return $null
}
