#Requires -Version 3
Param(
    [Parameter(Mandatory=$False,Position=1)]
    [string]$SubscriptionCsv,

    [Parameter(Mandatory=$False)]
    [string]$YoutubeDL

)


### Prerequisites ###
#####################

#Assume paths
$SubscriptionCsv = "$PSScriptRoot\$((Get-Item $PSCommandPath).BaseName).csv"
$YoutubeDL = "$PSScriptRoot\youtube-dl.exe"
$FFMpegLocation = "$PSScriptRoot"
$FFProbe = "$PSScriptRoot\ffprobe.exe"

#Check prereqs
ForEach ($PreReq in @($YoutubeDL, "$FFMpegLocation\FFMpeg.exe", $FFProbe)) {

    If (-Not(Test-Path $PreReq)) {
        Throw "$PreReq not found"
    }
    
}

#Process subscriptions
If (Test-Path $SubscriptionCsv) {

    $Subscriptions = Import-Csv $SubscriptionCsv -Delimiter ";"
    $Subscriptions | ForEach {

        #Define variables
        $Destination = $_.Destination
        $DRSlug = $_.DRSlug
        $SeriesDescriptionFile = "$Destination\_Description.txt"
        $SeriesDownloadLog = "$Destination\_DownloadLog.csv"
        $SeriesDownloadResult = @()

        #Get series
        Try {$Series = Invoke-RestMethod "http://www.dr.dk/mu/bundle?Slug=%22$($DRSlug)%22" -UseDefaultCredentials}
        Catch {Return}

        #Determine show title
        $Title = $Series.Data.Title
        Write-Verbose "Forsøger at hente senest tilgængelige episoder af ""$Title"" til ""$Destination""" -Verbose

        #Create destination if not exist
        If (-Not(Test-Path($Destination))) {
            Try {
                Write-Verbose "Destinationen ""$Destination"" findes ikke, opretter mappe" -Verbose
                New-Item $_.Destination -ItemType Directory | Out-Null
            }
            Catch {
                Write-Warning "Destinationen ""$Destination"" kunne ikke oprettes"
                Return
            }
        }

        #Output series description if not exist
        If (-Not(Test-Path($SeriesDescriptionFile))) {
            "$($Series.Data.Title)" | Out-File $SeriesDescriptionFile -Encoding default -Append
            "$($Series.Data.SubTitle)" | Out-File $SeriesDescriptionFile -Encoding default -Append
            "$($Series.Data.SubTitle)" | Out-File $SeriesDescriptionFile -Encoding default -Append
            "$($Series.Data.OnlineGenreText)" | Out-File $SeriesDescriptionFile -Encoding default -Append
            "DR Slug: $($Series.Data.Slug)" | Out-File $SeriesDescriptionFile -Encoding default -Append
        }

        #Get episodes
        #[int]$OptionalSeason = $_.OptionalSeason
        $BaseFileName = $_.BaseFileName
        #Handles an issue where a previous unrelated program's Slug has accidentally been added to a Series' bundle. DRSlug is the series slug, not the program' Slug.
        $Members = $Series.Data.Relations | Where Kind -eq "Member" | Where Slug -ne $DRSlug
        Try {$Episodes = ($Members | % {(Invoke-RestMethod "http://www.dr.dk/mu/programcard?Slug=%22$($_.Slug)%22").Data})}
        Catch {Return}
        <#
        $global:Index = 1
        $EpisodeList = $Episodes | 
            Select Title, PresentationUri, Slug, 
            @{N="AirTime";E={[datetime]$_.PrimaryBroadcastStartTime}}, 
            @{N="SeriesTitle";E={$series.data.title}}, 
            @{N="SeasonNo";E={$_.SeasonNumber}}, 
            @{N="EpisodeNo";E={$global:Index;$global:Index+=1}}
        #>

        <#
        #Cannot use this as DR sucks at assigning correct episode numbers!
        $EpisodeList = $Episodes | 
            Select Title, PresentationUri, Slug, 
            @{N="AirTime";E={[datetime]$_.PrimaryBroadcastStartTime}}, 
            @{N="SeriesTitle";E={$Series.data.title}}, 
            @{N="SeasonNo";E={$_.SeasonNumber}}, 
            @{N="EpisodeNo";E={$_.Production.PresentationEpisodeNumber}}
        #>

        
        #Just grabbing EpisodeNumber almost worked for Matador, but two episodes had no number specified
        $EpisodeList = $Episodes | 
            Select Title, PresentationUri, Slug, 
            @{N="AirTime";E={[datetime]$_.PrimaryBroadcastStartTime}}, 
            @{N="SeriesTitle";E={$Series.data.title}}, 
            @{N="SeasonNo";E={$_.SeasonNumber}}, 
            @{N="EpisodeNo";E={$_.EpisodeNumber}}
        

        <#
        #Workaround, counting episode number - doesn't always start numbering at 1
        $EpisodeList = @()
        $LastSeason = 0
        ForEach ($Episode in $Episodes) {

            #Determine episode number - need to count as DR sucks at assigning correct episode numbers to .Production.PresentationEpisodeNumber
            If ($Episode.SeasonNumber -ne $LastSeason) {
                $EpisodeNo = 1
            } Else {
               $EpisodeNo = $LastEpisode + 1
            }

            #Add entry to EpisodeList
            $EpisodeList += [PSCustomObject]@{
                Title = $Episode.Title;
                PresentationUri = $Episode.PresentationUri;
                Slug = $Episode.Slug;
                AirTime = [datetime]$Episode.PrimaryBroadcastStartTime;
                SeriesTitle = $Series.data.title;
                SeasonNo = $Episode.SeasonNumber;
                EpisodeNo = $EpisodeNo;
                
            }

            #Record last season and episode added to list
            $LastSeason = $Episode.SeasonNumber
            $LastEpisode = $EpisodeNo

        }
        #>

        #$EpisodeList | ft Title, SeasonNo, EpisodeNo -AutoSize #debug
        Write-Verbose "$($EpisodeList.Count) episode(r) listet på hjemmeside, $(@($EpisodeList | Where PresentationUri).Count) episode(r) tilgængelige for download" -Verbose

        #Get download log
        Remove-Variable DownloadLog -ErrorAction SilentlyContinue
        If (Test-Path $SeriesDownloadLog) {
            $DownloadLog = Import-Csv $SeriesDownloadLog -Delimiter ';'
        }

        #Match against downloaded episodes
        $DownloadList = @($EpisodeList | Where {$_.PresentationUri -and $_.Slug -notin ($DownloadLog | Where Success -eq 'TRUE' | Select -Expand Slug)})
        If ($DownloadList.Count -gt 0) {

            Write-Verbose "Henter $($DownloadList.Count) manglende episode(r)" -Verbose

            #Attempt to download next episode
            $DownloadList | % {

                #Define vars
                $DownloadDate = (Get-Date -Format s).Replace("T"," ")

                #Trying not to define filename, but derive from title
                #$FileName = "$BaseFileName - S$($_.SeasonNo.ToString("00"))E$($_.EpisodeNo.ToString("00")).%(ext)s" 
                $FileName = "$($_.Title).%(ext)s"

                <#
                If ($OptionalSeason) {
                    $FileName = "$BaseFileName - S$($OptionalSeason.ToString("00"))E$($_.EpisodeNo.ToString("00")).%(ext)s"
                } Else {
                    $FileName = "$BaseFileName - E$($_.EpisodeNo.ToString("00")).%(ext)s"
                }
                #>
                    
                #Attempt download
                #Not using "--convert-subtitles srt" and "--embed-subs", as DR's VTT subs don't seem to be convertible to SRT by FFmpeg/youtube-dl
                Write-Verbose "Downloader ""$($_.Title)"" til ""$($Destination)""" -Verbose
                Push-Location $Destination
                
                
                #& $YoutubeDL --no-progress --no-overwrites --no-warnings --output ""$FileName"" --format mp4 --write-thumbnail --write-description --add-metadata --all-subs $_.PresentationUri
                #& $YoutubeDL --no-progress --no-overwrites --no-warnings --output ""$FileName"" --write-thumbnail --write-description --add-metadata --all-subs $_.PresentationUri
                #& $YoutubeDL --no-warnings --write-thumbnail --write-description --add-metadata --all-subs --console-title $_.PresentationUri --ffmpeg-location ""$FFMpegLocation""
                & $YoutubeDL --no-warnings --output ""$FileName"" --write-thumbnail --write-description --add-metadata --all-subs --console-title $_.PresentationUri --ffmpeg-location ""$FFMpegLocation""
                
                Write-Verbose "Hentning fuldført med returkode $LASTEXITCODE" -Verbose
                Pop-Location

                #Log result
                $SeriesDownloadResult += [pscustomobject]@{
                    Slug = $_.Slug;
                    Date = $DownloadDate;
                    Success = [bool]($LASTEXITCODE -eq 0);
                    ExitCode = $LASTEXITCODE;
                }

            }

            #Append download results to log
            $SeriesDownloadResult | Export-Csv $SeriesDownloadLog -NoTypeInformation -Delimiter ';' -Append

        } Else {

            Write-Verbose "Ingen tilgængelige manglende episoder" -Verbose

        }

    }

    Start-Sleep 5

} Else {

    Write-Warning "Abonnementsliste ""$SubscriptionCsv"" ikke fundet"
    Break

}
