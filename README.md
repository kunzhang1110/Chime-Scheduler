# Hourly Chime

Hourly Chime plays a sound on Winows at a certain interval.

## How to use it

1. Go to Task Scheduler
1. Select Action - Create Task - Type a name
1. Go to Triggers and select New...
   - In settings select Daily,
   - In advanced settings, check Repeat task every 1 hour for a duration of 12 hours.
   - Click Ok.
1. Go to Actions and select New...
   - In Settings - Program/script, type `powershell.exe`
   - In Settings - Add arguments, type `-WindowStyle hidden -File C:\path\to\HourlyChime.ps1`
     - `-WindowStyle hidden` hides the powershell window
1. Go to Settings
   - Check `Run task as soon as possible after a scheduled start is missed
`

## HourlyChime.ps1

The ps1 script plays a sound file locating at 'C:\Windows\Media\Alarm03.wav'.
