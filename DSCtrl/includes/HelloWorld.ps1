function Invoke-HelloWorld{
    $Global:SyncHash.print("Hello World!", $false)
    $Global:SyncHash.MessageBox.Show("Name:")

    $Global:SyncHash.print(("Distinguished Name: " + $computer.distinguishedname), $false)
    $Global:SyncHash.print(("Name: " + $computer.name), $false)

}
