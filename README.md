# slack-dm-fetch
Fetch Slack DM history and export to CSV

## Setup
`pip install -r requirements.txt`

## Usage
`env SLACK_API_TOKEN=<token> python fetch.py --user <username to find DMs with>`

### Example output
```
Found 55 messages with user chardizzeroony
Saving to chardizzeroony-message-history.csv...
Save complete
```
