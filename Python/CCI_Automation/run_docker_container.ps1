# Build the Docker image
docker build -t test1 .

# Get the IP address for DISPLAY variable
$ipAddress = ipconfig | Select-String -Pattern 'IPv4' | ForEach-Object { 
    if ($_.Line -match '(\d{1,3}\.){3}\d{1,3}') { 
        [PSCustomObject]@{IPAddress = $matches[0]}
    }
} | Select-Object -Last 1

# print the IP address
# set-variable -name display -value 172.26.48.1:0.0


# Extract the IP address from the object
$ipAddress = $ipAddress.IPAddress.Trim()
Write-Host "Extracted IP Object: $ipAddress"
# Append ":0.0" to the IP address
$display = "{0}:0.0" -f $ipAddress

# Print the result
Write-Host "IP Address: $display"

# Get the user's Downloads and Documents folder paths
$downloadsPath = [System.Environment]::GetFolderPath('UserProfile') + "\Downloads"
$documentsPath = [System.Environment]::GetFolderPath('MyDocuments')

# Run the Docker container with dynamic paths
$xAuthority = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('UserProfile'), '.Xauthority')
docker run -ti --rm -e DISPLAY=$display -e XAUTHORITY=$xAuthority -v "${downloadsPath}:/app/Downloads" -v "${documentsPath}:/app/Documents" test1

# Pause and wait for user input
Read-Host -Prompt "Press Enter to exit"