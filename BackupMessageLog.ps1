#---------------------------------------------------
# Powerscript Shell
# Auto Backup & Copy Files
#---------------------------------------------------

#Function 7zip dengan paramater
Function CopyMessageLog
{
    Param
    (
        [string]$strFolder   # Folder tempat Message Log ()
    )
    # Ambil Tanggal Saat Ini
    $Now = Get-Date

    #Konversi tanggal ke string dengan format tanggal YYYY_MM_DD
    $TxtFileDate = $Now.ToString("yyyy_MM_dd")

    $strSourceFile = $strFolder +  "\Message.log"
    $strTargetFile = $strFolder +  "\BackupLog\Message_" + $TxtFileDate + ".log"

    If(Test-Path $strSourceFile)
    {
        # Copy Message Log
        Copy-Item -Path "$strSourceFile" -Destination $strTargetFile

        # Hapus Message Log
        Remove-item -Path $strSourceFile
    }

    # Kosongkan variabel
    $TxtFileDate = ""
    $strSourceFile = ""
    $strTargetFile = ""
}

Function CopyErrorLog
{
    Param
    (
        [string]$strFolder   # Folder tempat Error Log 
    )
    # Ambil Tanggal Saat Ini
    $Now = Get-Date

    #Konversi tanggal ke string dengan format tanggal YYYY_MM_DD
    $TxtFileDate = $Now.ToString("yyyy_MM_dd")

    $strSourceFile = $strFolder +  "\error.log"
    $strTargetFile = $strFolder +  "\BackupLog\error_" + $TxtFileDate + ".log"

    If(Test-Path $strSourceFile)
    {
        # Copy Message Log
        Copy-Item -Path "$strSourceFile" -Destination $strTargetFile

        # Hapus Message Log
        Remove-item -Path $strSourceFile
    }

    # Kosongkan variabel
    $TxtFileDate = ""
    $strSourceFile = ""
    $strTargetFile = ""
}

CopyMessageLog "C:\fb-plants\debug"
CopyErrorLog "C:\fb-plants\debug"


