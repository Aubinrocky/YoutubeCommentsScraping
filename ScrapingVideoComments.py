# Import of libraries
import os
import googleapiclient.discovery

# Access to Google Sheet:
from oauth2client.service_account import ServiceAccountCredentials
import gspread
scope = ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    r"C:\Users\aubin\OneDrive\Documents\DataCode\SheetsAPIkey.json", scope)
client = gspread.authorize(creds)
sheet = client.open("YoutubeComments").sheet1


# Youtube video ID :
VideoID = "ezBnVjmsTP8"

# Access to Youtube Url :
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
api_service_name = "youtube"
api_version = "v3"
DEVELOPER_KEY = "AIzaSyD2_jrLtJ51j01YETaBrSmDFBrKcPcMx80"

# Collection of Data
youtube = googleapiclient.discovery.build(
    api_service_name, api_version, developerKey=DEVELOPER_KEY)
request = youtube.commentThreads().list(
    part="snippet",
    order="time",
    videoId=VideoID,
    maxResults='100')
response = request.execute()

# Cleaning architecture of the dataset
header = ['VideoId', 'channelId', 'Name',
          'Comment', 'Time', 'Likes', 'Reply Count']
data = []
for i in response["items"]:
    videoId = i["snippet"]['topLevelComment']["snippet"]["videoId"]
    authorChannelUrl = i["snippet"]['topLevelComment']["snippet"]["authorChannelUrl"]
    author = i["snippet"]['topLevelComment']["snippet"]["authorDisplayName"]
    comment = i["snippet"]['topLevelComment']["snippet"]["textDisplay"]
    published_at = i["snippet"]['topLevelComment']["snippet"]['publishedAt']
    likes = i["snippet"]['topLevelComment']["snippet"]['likeCount']
    replies = i["snippet"]['totalReplyCount']
    data.append([videoId, authorChannelUrl, author, comment,
                 published_at, likes, replies])


# Insert data into the Google Sheet :
no_Rows = len(data)
sheet.update('A1:G1', [header])
sheet.format('A1:G1', {'textFormat': {'bold': True}})
sheet.update('A2:G'+str(no_Rows+1), data)
