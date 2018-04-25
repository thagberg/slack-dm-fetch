import os
import csv
import argparse
from collections import defaultdict
from slackclient import SlackClient

USER_MAP = defaultdict(lambda: 'You')

parser = argparse.ArgumentParser(
    description="Fetch Slack DM conversation and export to spreadsheet")
required_args = parser.add_argument_group('required')
required_args.add_argument('--user', 
                           dest='user', 
                           action='store', 
                           required=True, 
                           help='Name of user to fetch DMs with')
args = parser.parse_args()

slack_token = os.environ['SLACK_API_TOKEN']
client = SlackClient(slack_token)

dm_channels = client.api_call(
    "im.list"
)

users = client.api_call(
    "users.list"
)

find_user_name = args.user
find_user_id = None
for user in users['members']:
    if user['name'] == find_user_name:
        find_user_id = user['id']
        break
USER_MAP[find_user_id] = find_user_name

if not find_user_id:
    raise Exception("Couldn't find user: {}".format(find_user_name))

found_dm_id = None
for dm_channel in dm_channels['ims']:
    if dm_channel['user'] == find_user_id:
        found_dm_id = dm_channel['id']
        break

if not found_dm_id:
    raise Exception("No DMs with {}".format(find_user_name))

more_messages = True
saved_messages = []
last_ts = 0
while more_messages:
    messages = client.api_call(
        "im.history",
        channel=found_dm_id,
        oldest=last_ts
    )
    saved_messages += messages['messages']
    more_messages = messages['has_more']
    if more_messages:
        last_ts = messages['latest']

filename = '{}-message-history.csv'.format(find_user_name)
print("Found {} messages with user {}".format(len(saved_messages), find_user_name))
print("Saving to {}...".format(filename))

with open(filename, 'w', newline='') as csvfile:
    csvwriter = csv.writer(csvfile, 
                           dialect='excel', 
                           delimiter=',', 
                           quoting=csv.QUOTE_MINIMAL)
    csvwriter.writerow(['Time', 'Author', 'Content', 'Is Starred'])
    for message in saved_messages:
        username = USER_MAP[message['user']]
        text = message['text']
        time = message['ts']
        starred = 'yes' if 'is_starred' in message else ''
        csvwriter.writerow([time, username, text, starred])

print("Save complete")