from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import pandas as pd
import os


# ----------------------------------------
# CONFIGURATION
# ----------------------------------------
sharepoint_site = ""
document_library = ""
excel_filename = ""

username = "" # Suggest an Env.py or Config.py file
password = ""


# ----------------------------------------
# CONNECT TO SHAREPOINT
# ----------------------------------------
try:
    ctx = ClientContext(sharepoint_site).with_credentials(
        UserCredential(username, password)
    )
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    site_title = web.properties.get("Title", sharepoint_site)
    print(f"Successfully connected to SharePoint site: {site_title}")
except Exception as e:
    print(f"Failed to connect to SharePoint: {e}")
    sys.exit(1)

# File path on SharePoint
file_url = f"/sites/YourSite/{document_library}/{excel_filename}"

print(file_url)