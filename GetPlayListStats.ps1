param(
    [Parameter(Mandatory)]
    $GoogleApiKey
)

function Get-YouTubePlaylist {
    param(
        [Parameter(Mandatory)]
        [string]
        $playListId
    )

    do {
        $URL = "https://www.googleapis.com/youtube/v3/playlistItems?part=snippet&playlistId={0}&maxResults=50&key={1}&pageToken={2}" -f $playListId, $GoogleApiKey, $nextPageToken
        $r = Invoke-RestMethod $URL
        $nextPageToken = $r.nextPageToken
        $r.items.snippet
    } until ($null -eq $nextPageToken)

}

function Get-YouTubeVideo {
    param(
        [Parameter(ValueFromPipeline)]
        $YouTubeVideo
    )

    Process {
        $videoId = $YouTubeVideo.resourceId.videoId
        Write-Progress -Activity "Getting YouTube Stats - $($YouTubeVideo.channelTitle)" -Status "Processing video stats - $($YouTubeVideo.title)"
        $URL = "https://www.googleapis.com/youtube/v3/videos?id={0}&key={1}&part=snippet,contentDetails,statistics,status" -f $videoId, $GoogleApiKey
        
        $r = Invoke-RestMethod -uri $URL

        if ($r.items.count -gt 0) {
            $publishedAt = $r.items.snippet.publishedAt
        
            $publishedAt = (Get-Date $publishedAt).ToString("yyyy-MM-dd HH:mm:ss")
        
        
            $stats = $r.items.statistics
            [pscustomobject][Ordered]@{
                Published     = $publishedAt
                Year          = (Get-Date $publishedAt).Year
                Month         = (Get-Date $publishedAt).Month
                MonthName     = (Get-Date $publishedAt).ToString("MMM")
                Title         = $r.items.snippet.title
                ViewCount     = $stats.viewCount
                LikeCount     = $stats.likeCount
                DislikeCount  = $stats.dislikeCount
                FavoriteCount = $stats.favoriteCount
                CommentCount  = $stats.commentCount
                Url           = 'https://www.youtube.com/watch?v={0}' -f $videoId
            }
        }
    }    
}

# $playLists = ConvertFrom-Csv @"
# FileName,ID
# NYC PS Meetup, PL5uoqS92stXiRX67A85FyrXvtn71eTiWO
# Maggie Shorts, PL8xWO_v7SZjsvD6WjCwuMGXb3kdBo_SiX
# "@

$playLists = Import-Csv ./playlists.csv

$playLists | ForEach-Object {
    $fileName = "{0}-{1}.xlsx" -f $_.FileName, $_.ID

    $workSheetName = (Get-Date).ToString("yyyyMMddHHmmss - MMM dd yyyy") 

    Get-YouTubePlaylist $_.ID |
    Get-YouTubeVideo |
    Sort-Object year, month | 
    Export-Excel -Path "./$fileName" -WorksheetName $workSheetName -AutoSize -AutoFilter
}