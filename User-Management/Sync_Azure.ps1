	Write-Host "Initializing Azure AD Delta Sync..." -ForegroundColor Yellow

	Start-ADSyncSyncCycle -PolicyType Delta

	#Wait 10 seconds for the sync connector to wake up. 
	Start-Sleep -Seconds 10

	#Display a progress indicator and hold up the rest of the script while the sync completes.
	While(Get-ADSyncConnectorRunStatus)
	{
		Write-Host "." -NoNewline
		Start-Sleep -Seconds 10
	}

Write-Host " | Complete!" -ForegroundColor Green
