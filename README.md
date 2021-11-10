# Use PowerShell to Track YouTube Video Statistics

Track the number of views, likes, and comments on a YouTube video, or playlist and update and Excel spreadsheet with the latest results. Plus, it runs every day on a schedule, and checks the spreadsheet back into the repository.

## Technology used:

    - PowerShell Excel module (`ImportExcel.psm1`)
    - YouTube API
    - GitHub Actions

## Steps:

1. Fork the repository
1. Get a YouTube API key https://developers.google.com/youtube/v3/getting-started
1. Create a `secret` for the YouTube API Key on the GitHub Repo - [Create a Secret](https://github.com/Azure/actions-workflow-samples/blob/master/assets/create-secrets-for-GitHub-workflows.md#:~:text=Creating%20secrets%201%20On%20GitHub%2C%20navigate%20to%20the,value%20for%20your%20secret.%207%20Click%20Add%20secret.)
    - `GoogleApiKey` is the name of the secret
1. Update the `playlists.csv` file with the YouTube playlist IDs you want to track
1. Run the GitHub Action workflow.

After it runs, it will create a new Excel file, or update it with a new sheet, and check the Excel xlsx back into the repository.

You can click on the Excel file, then click `view raw` to download it to your local machine to view.