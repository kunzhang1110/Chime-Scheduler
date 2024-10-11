$soundFile = 'C:\Windows\Media\Alarm03.wav'
$player = New-Object System.Media.SoundPlayer
$player.SoundLocation = $soundFile
$player.PlaySync()
Write-Output "Chime"



