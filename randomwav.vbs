Dim fso, folder, files, wavFileCount, wavFiles()

Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("C:\Temp")
Set files = folder.Files
wavFileCount = 0

' Get the list of all the wav files
For Each file in files
    name = lcase(file)
    wscript.echo "Found: " & name
    if right(name, 3) = "wav" then
        wavFileCount = wavFileCount + 1
        ReDim Preserve wavFiles(wavFileCount)
        wavFiles(wavFileCount - 1) = name
    end if
Next

' Now that we have just WAV files, figure out which one to play
If wavFileCount < 1 then
    wscript.echo "No wav files found, giving up"
    wscript.quit
End If

' Play the random WAV file
Randomize
randomIndex = Int(rnd * wavFileCount)
randomWavFile = wavFiles(randomIndex)
wscript.echo "Playing " & randomWavFile
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "start """ & randomWavFile & """"