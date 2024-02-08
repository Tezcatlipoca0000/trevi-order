"""
Program that takes the order to the provider, creates a new table with the used articles and send it to the provider
"""

# ************************************
# Sending imports
import os.path
import base64
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.message import EmailMessage

SCOPES = [
	"https://www.googleapis.com/auth/gmail.modify"
]

# ************************************
# Creating the order list imports
import pandas as pd

# ************************************
# Sending logic
def send():
	creds = None

	if os.path.exists("token.json"):
		creds = Credentials.from_authorized_user_file("token.json", SCOPES)

	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
			creds = flow.run_local_server(port=0)
		with open("token.json", "w") as token:
			token.write(creds.to_json()) 

	try:
		service = build("gmail", "v1", credentials=creds)
		message = EmailMessage()

		message.set_content("Cliente: ANTONIO OLGUIN CORREA\nDirección: FUENTE REAL #300 FUENTES DE ANAHUAC SAN NICOLAS DE LOS GARZA C.P. 66444 Tel: 8183833114\nPEDIDO ADJUNTO")
		message["To"] = "superfuentes300@gmail.com"
		message["From"] = "superfuentes300@gmail.com"
		message["Subject"] = "Pedido Súper Fuentes"
		message["CC"] = "superfuentes300@gmail.com"

		attachment = "pedido.xlsx"
		filename = os.path.basename(attachment)
		with open(attachment, "rb") as attch:
			attachment_data = attch.read()
		message.add_attachment(attachment_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)

		encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
		create_message = {"raw": encoded_message}
		send_message = (
			service.users()
			.messages()
			.send(userId="me", body=create_message)
			.execute()
		)

	except HttpError as error:
		print(f"an HttpError occurred: {error}")
		send_message = None

	return send_message

# ************************************
# Creating the order list logic
def main():
	xl_df = pd.read_excel("Provedores Todos.xlsm", sheet_name="Pedido", usecols="B:D,F")
	xl_df = xl_df[xl_df["Provedor"] == "Treviño"]
	xl_df = xl_df[xl_df["Pedido"].notnull()]
	xl_df = xl_df.drop("Provedor", axis=1)
	xl_df = xl_df.rename(columns={"Pedido": "Cantidad"})
	xl_df.to_excel("pedido.xlsx", index=False)

	send()

# ************************************
if __name__ == "__main__":
	main()