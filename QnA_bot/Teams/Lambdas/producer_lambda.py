import requests
import json
from msal import PublicClientApplication
import boto3
from botocore.exceptions import ClientError
from bs4 import BeautifulSoup

class TokenIntilialzer:
    def __init__(self, secret_name, region="us-east-1"):
        self.secret_name = secret_name
        self.secret_manager_client  = boto3.client(
            "secretsmanager",region_name = "us-east-1"
        )

    def get_secret_values_generate_token(self):
        get_secret_value_response = self.secret_manager_client.get_secret_value(SecretId = self.secret_name)
        client_id = get_secret_value_response["client_id"]
        username = get_secret_value_response["user_name"]
        password = get_secret_value_response["password"]
        tenant_id = get_secret_value_response["tenant_id"]

        authority = f"https://login.microsoftonline.com/{tenant_id}"
        scope = ["ChannelMessage.Read.All"]
        app=msal.PublicClientApplication(
        client_id = client_id,
        authority = authority
        )
        
        self.access_token = app.acquire_token_by_username_password(
            username=username,
            password=password,
            scope=scope
        )
class TeamsHandler(TokenIntilialzer):
    def __init__(self, teams_id,channel_id, channel_name, region="us-east-1"):
        self.teams_id =teams_id
        self.channel_id =channel_id
        self.channel_name =channel_name
        self.channel_message_url =f"https://graph.microsoft.com/v1.0/teams/{teams_id}/channels/{channel_id}/messages"
    def get_channel_messages(self):
        channel_message_response = requests.get(self.channel_message_url,headers={"Authorization":f"Bearer {self.access_token}"})
        for each in channel_message_response.json()["value"]:
            soup=None
            plain_text = ""
            try:
                soup=BeautifulSoup(each["body"]["content"],"html.parser")
                plain_text = soup.get_text(separator=" ")
            except:
                plain_text=each["body"]["content"]
            message = []
            message.append([
                each["id"],
                each["subject"],
                self.channel_name,
                each["createdDateTime"],
                each["summary"],
                each["importance"],
                plain_text.replace("\n",""),
                each["body"]["content"]

            ])
        self.channel_message  = message
class QueueHandler(TeamsHandler):
    def __init__(self, table_name, teams_id, channel_name,channel_id, queue_url, region="us-east-1"):
        super().__init__(
            teams_id=teams_id,
            channel_name=channel_name,
            channel_id  = channel_id
        )
        self.dynamodb = boto3.resource("dynamodb",region_name = region)
        self.channel_name = channel_name
        self.queue_url = queue_url
        self.table_name = table_name
    def check_if_message_pushed(self, message_id):
        table = self.dynamodb.Table(self.table_name)
        try:
            response = table.get_item(Key = {"id":message_id})
            item = response.get("Item")
            if item:
                return True
            else:
                return False
        except ClientError as e:
            return False
    def psuh_message_to_sqs(self, message):
        sqs=boto3.client("sqs",region_name = "us-east-1")
        message_body  = json.dump(message)
        response = sqs.send_message(
            QueueUrl = self.queue_url,
            MessageBody = message_body,
        )
    def udpdate_table_with_messages(self):
        table = self.dynamodb.Table(self.table_name)
        for each in self.channel_message:
            if not self.check_if_message_pushed(each[0]) and each[6] !="":
                self.psuh_message_to_sqs(
                    {
                        "message_id":each[0],
                        "subject":each[1],
                        "summary":each[4],
                        "content":each[6]
                    }
                )
            response = table.update_item(
                Key = {"id":each[0]},
                UpdateExpression = "SET #subject = :subject,#createdDateTime = :createdDateTime,#summary = :summary,#content = :content,#content_unfiltered = :content_unfiltered,#channel_name = :channel_name",
                ExpressionAttributeNames = {
                    "#subject":"subject",
                                   "#createdDateTime":"createdDateTime",
                    "#summary":"summary",
                    "#content":"content",
                    "#content_unfiltered":"content_unfiltered",
                    "#channel_name":"channel_name",
     
     
                },
                ExpressionAttributeValues = {
                    ":subject":each[1],
                    ":createdDateTime":each[3],
                    ":channel_name":each[2],
                    ":summary":each[4],
                    ":content_unfiltered":each[7],
                    ":content":each[6]
                }
                ReturnValues = "UPDATE_NEW"
            )
def lambda_handler(event,context):
    handler = QueueHandler(
        table_name= event["team_name"],
        teams_id = event["team_id"],
        channel_id = event["channel_id"],
        channel_name= event["channel_name"]
    )
    handler.get_secret_values_generate_token()
    handler.get_channel_messages()
    handler.udpdate_table_with_messages()