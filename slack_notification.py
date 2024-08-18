import os
from slack import WebClient
from slack.errors import SlackApiError

client = WebClient(token='xoxb-797224974471-7437246651842-hK9DIM4fr85QOtkpJxEA6OfZ')

# def invite_bot(channel_id, bot_user_id):
#     try:
#         response = client.conversations_invite(
#             channel = channel_id,
#             users = bot_user_id
#         )
#         if response['ok']:
#             print(f"Bot invited to the channel: {channel_id}")
#         else:
#             print(f"Failed to invite bot: {response['error']}")
#     except SlackApiError as e:
#         print(f"Got an error: {e.response['error']}")


def send_message(channel_id):    
    try:
        response = client.chat_postMessage(
            channel= channel_id,
            text="I will give updates about monthly deviations here in this channel")
        assert response["message"]["text"] != None
    except SlackApiError as e:
        # You will get a SlackApiError if "ok" is False
        assert e.response["ok"] is False
        assert e.response["error"]  # str like 'invalid_auth', 'channel_not_found'
        print(f"Got an error: {e.response['error']}")


channel_id = 'C07CNBLK2B0'
# bot_user_id = 'U07CV78K5QS'
# invite_bot(channel_id, bot_user_id)
send_message(channel_id)