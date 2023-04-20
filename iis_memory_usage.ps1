$W3WPProcesses = Get-WmiObject -Query "SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name LIKE '%w3wp%'"
$Instances = $W3WPProcesses | ForEach-Object {
    $CommandLine = $_.CommandLine
    $AppPoolNameStart = $CommandLine.IndexOf('-ap "') + 5
    $AppPoolNameEnd = $CommandLine.IndexOf('"', $AppPoolNameStart)
    $AppPoolName = $CommandLine.Substring($AppPoolNameStart, $AppPoolNameEnd - $AppPoolNameStart)
    [PSCustomObject]@{
        mypid = $_.ProcessId
        appPoolName = $AppPoolName
    }
} | Sort-Object -Property mypid

$InstancesWithNumbers = $Instances | Group-Object -Property appPoolName | ForEach-Object {
    $AppPoolName = $_.Name
    $Processes = $_.Group
    $Counter = 1

    $Processes | ForEach-Object {
        [PSCustomObject]@{
            mypid = $_.mypid
            appPoolName = $AppPoolName
            num = $Counter
        }
        $Counter++
    }
} | Sort-Object -Property mypid

$result = $InstancesWithNumbers | ForEach-Object {
    $Process = Get-Process -Id $_.mypid
    [PSCustomObject]@{
        mypid = $_.mypid
        num = $_.num
        appPoolName = $_.appPoolName
        cpuUsage = $Process.CPU
        memoryAllocated = $Process.PrivateMemorySize64
        memoryUsed = $Process.WorkingSet64
        threadsCount = $Process.Threads.Count
    }
} | Sort-Object -Property mypid | ConvertTo-Json

$result
