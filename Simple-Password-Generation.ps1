$max = 25 # Set how many passwords you want generated here
$passwords = @() # Variable all the passwords are saved to

foreach ($count in 1..$($max)){
    $UserPassword = curl -Uri http://www.dinopass.com/password/simple # Gets a simple password from Dinopass.com using their API
    $passwords += $($($UserPassword.Content).substring(0,1).Toupper()+$($UserPassword.Content).substring(1)) # This capitalizes the first letter of the password
}

# The following line saves the password to the user's desktop as a CSV file
$passwords | Select-Object @{Name='Password';Expression={$_}} | Export-Csv "$([Environment]::GetFolderPath("Desktop"))\GeneratedPasswords.csv" -Force
