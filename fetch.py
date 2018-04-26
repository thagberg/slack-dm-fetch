import os
import csv
import argparse
import time
import sys
import datetime
from collections import defaultdict
from slackclient import SlackClient

# const value for how many messages we want to fetch
# in a single "page"
PAGE_COUNT = 100
# mapping of User IDs to User display names
# a defaultdict is like a dict/dictionary, except
# if you try to look up a key that doesn't exist, it
# uses the "constructor" function you provide as an argument
# to create a default value for that key
# lambda: 'You' is basically an allowed "constructor" for
# creating the string "You".  This way, if we look up
# usernames and don't find the key, that means it was "You"
# in the conversation
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


def fetch_messages(channel_id, latest_ts=None):
    """Fetch messages from channel/conversation; supports paging

    channel_id -- Slack channel/conversation ID
    latest_ts -- Upper bound timestamp for paging
    """
    messages = {} 
    if latest_ts:
        messages = client.api_call(
            "im.history",
            channel=channel_id,
            latest=latest_ts,
            count=PAGE_COUNT
        )
    else:
        messages = client.api_call(
            "im.history",
            channel=channel_id,
            count=PAGE_COUNT
        )
    return messages


def find_bounds(messages):
    """Find oldest and latest message timestamps

    messages -- List of Slack messages to check
    """
    oldest=sys.float_info.max
    latest=0
    for message in messages:
        ts = float(message['ts'])
        if ts > latest:
            latest = ts
        if ts < oldest:
            oldest = ts
    return (oldest, latest)


if __name__ == "__main__":
    # First we need to get a list of every direct message
    # "channel" that our user has access to
    dm_channels = client.api_call(
        "im.list"
    )

    # Then we get the list of all users in this workspace
    # This is because on the command-line we provided a user
    # NAME, and we need to find the appropriate user ID
    users = client.api_call(
        "users.list"
    )

    # Given the username from the command-line arguments,
    # search the users we fetched from slack to find the
    # correct user ID
    find_user_name = args.user
    find_user_id = None
    for user in users['members']:
        if user['name'] == find_user_name:
            find_user_id = user['id']
            break
    USER_MAP[find_user_id] = find_user_name

    if not find_user_id:
        raise Exception("Couldn't find user: {}".format(find_user_name))

    # Now find the channel ID of the conversation "channel"
    # we have with this user.
    found_dm_id = None
    for dm_channel in dm_channels['ims']:
        if dm_channel['user'] == find_user_id:
            found_dm_id = dm_channel['id']
            break

    if not found_dm_id:
        raise Exception("No DMs with {}".format(find_user_name))

    # First we make one attempt at fetching the messages from
    # the conversation we located previously.
    # If there are more messages available (identified by the
    # "has_more" field) then we keep fetching more messages,
    # using a technique called "paging" where we give it a new
    # set of "bounds" each time we ask for more messages.  This way,
    # we can eventually fetch all the messages even though we can only
    # fetch them at a certain number per request
    more_messages = True
    saved_messages = []
    messages = fetch_messages(found_dm_id)
    more_messages = messages['has_more']
    saved_messages = saved_messages + messages['messages']
    (oldest_ts, last_ts) = find_bounds(messages['messages'])
    while more_messages:
        messages = fetch_messages(found_dm_id, oldest_ts)
        (oldest_ts, last_ts) = find_bounds(messages['messages'])
        more_messages = messages['has_more'] or len(messages['messages']) >= PAGE_COUNT
        saved_messages += messages['messages']

    # sort messages in ascending chronological order
    saved_messages = sorted(saved_messages, key=lambda message: message['ts'])

    filename = '{}-message-history.csv'.format(find_user_name)
    print("Found {} messages with user {}".format(len(saved_messages), find_user_name))
    print("Saving to {}...".format(filename))

    # We now open a new CSV file, which we'll begin writing to
    # The "with" mechanism is a cool feature in Python which lets
    # you open a file or some other resource and let Python automatically
    # close it when it goes out of scope
    with open(filename, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile, 
                            dialect='excel', 
                            delimiter=',', 
                            quoting=csv.QUOTE_MINIMAL)
        # Here we're adding a "header" row to the CSV
        csvwriter.writerow(['Time', 'Author', 'Content', 'Is Starred'])
        # Now we're iterating over all the messages we found
        # For each one, we're going to do a little bit of formatting,
        # and then we're going to write it as a new row in the CSV file
        for message in saved_messages:
            username = USER_MAP[message['user']]
            text = message['text']
            time = datetime.datetime.fromtimestamp(
                float(message['ts'])
            ).strftime("%Y-%m-%d %H:%M:%S")
            starred = 'yes' if 'is_starred' in message else ''
            csvwriter.writerow([time, username, text, starred])

    print("Save complete")