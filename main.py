from easysms import EasySMSClient
import dotenv
import os

dotenv.load_dotenv()

client = EasySMSClient(api_key=os.getenv("API_KEY"))