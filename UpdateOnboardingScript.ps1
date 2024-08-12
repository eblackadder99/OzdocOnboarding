## This needs to be placed in C:\Ozdoc\Update

## Get the latest release information from GitHub
$latestRelease = Invoke-RestMethod -Uri 'https://api.github.com/repos/eblackadder99/OzdocOnboarding/releases/latest'
$latestTag = $latestRelease.tag_name
$downloadUrl = "https://github.com/eblackadder99/OzdocOnboarding/archive/refs/tags/$latestTag.zip"

## Get the current version of the script
$updateCheck = "C:\Ozdoc\UserOnboarding\Update\version.txt"
$updateCheckContent = (Get-Content -Path $updateCheck -Raw).Trim()

## Download the latest release zip file
Invoke-WebRequest -Uri $downloadUrl -OutFile C:\Ozdoc\UserOnboarding.zip

## Extract the downloaded zip file
Expand-Archive -Path C:\Ozdoc\UserOnboarding.zip -DestinationPath C:\Ozdoc\ -ErrorAction SilentlyContinue

## Get the name of the extracted folder
$extractedFolder = Get-ChildItem -Path C:\Ozdoc\ | Where-Object { $_.PSIsContainer -and $_.Name -like "OzdocOnboarding-*" } | Select-Object -First 1

## Remove the existing 'UserOnboarding' folder if it exists
if (Test-Path -Path C:\Ozdoc\UserOnboarding) {
    Remove-Item -Path C:\Ozdoc\UserOnboarding -Recurse -Force
}

## Rename the extracted folder to 'UserOnboarding'
Rename-Item -Path $extractedFolder.FullName -NewName C:\Ozdoc\UserOnboarding -Force

## Remove the downloaded zip file
Remove-Item -Path C:\Ozdoc\UserOnboarding.zip -Force

## Get the latest version of the script afer updating
$updatedCheck = "C:\Ozdoc\UserOnboarding\Update\version.txt"
$updatedCheckContent = (Get-Content -Path $updatedCheck -Raw).Trim()

## Check that update was successful
if (-NOT($updateCheckContent -eq $updatedCheckContent)) {
    Write-Output `n "Update was successful, restarting user creation script...`r"
    Start-Sleep -Seconds 2
    $latestScript = Get-ChildItem -Path C:\Ozdoc\UserOnboarding\ | Where-Object { $_.Name -like "UserOnboarding-*.ps1" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1    
    & $latestScript.FullName
    exit
}
    else {
        Write-Host "`nUpdate was not successful.`r"
        Start-Sleep -Seconds 3
        Write-Host "Press Enter to exit..."
        [void][System.Console]::ReadLine()
        exit
    }