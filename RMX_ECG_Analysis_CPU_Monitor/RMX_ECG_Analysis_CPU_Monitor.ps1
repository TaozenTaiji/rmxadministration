While($true)
{
    $cpuutil=(get-counter -Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -MaxSamples 300 <#Length of sample time in seconds#> |
        Select-Object -ExpandProperty countersamples | Select-Object -ExpandProperty cookedvalue | Measure-Object -Average).average
    out-file -InputObject $cpuutil -FilePath 'C:\Users\zhilliker\OneDrive - Rhythmedix LLC\Documents\WindowsPowerShell\Scripts\CPUTest.txt' -Append  <#for testing as a background job#>

    If ($cpuutil -le 5 <#minimum CPU threshhold#> -and $AnalysisJobs -gt 0 <# Test for messages from service is >0 #>)
    {
        Restart-Computer
    }

    Start-sleep 30
}