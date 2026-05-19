from easysms import EasySMSClient
from ui.config import get_api_key

api_key = get_api_key()
if not api_key:
    raise RuntimeError("API key is not configured. Set it in app Settings.")

client = EasySMSClient(api_key=api_key)