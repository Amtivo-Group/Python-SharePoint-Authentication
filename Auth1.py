from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import pandas as pd
import os
import sys


print(F"Config Started.")
# ----------------------------------------
# CONFIGURATION
# ----------------------------------------
sharepoint_site = "https://britishassessmentbureau.sharepoint.com/sites/FPA"
document_library = ""
excel_filename = ""

username = os.getenv('userEmail') # Suggest an Env.py or Config.py file
password = os.getenv('userPassword')

print(F"Config Complete. Trying to connect")


# ----------------------------------------
# CONNECT TO SHAREPOINT
# ----------------------------------------
try:
    print("Trying to connect.......")
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

