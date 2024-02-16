import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1bfmsfL02MqCkmuoLs-LDBIIUtOKuP5Z8spS4fZC_kdA"
SAMPLE_RANGE_NAME = "engenharia_de_software!A3:I28"
RANGE_SCHOOL_ABSENCES = "engenharia_de_software!C4:C27"


def main():

  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    school_absences = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=RANGE_SCHOOL_ABSENCES)
        .execute()
    )
    values = result.get("values", [])
    abscenses = school_absences.get("values", [])
    

  
    final_grade = []

    for  i,row in enumerate(values):
        if i > 0:
           row_grade_1 = int(row[3])
           row_grade_2 = int(row[4]);
           row_grade_3 = int(row[5]);
           nota_final = (row_grade_1 + row_grade_2 + row_grade_3)/30
           
           nota_final = round(nota_final, 2)
           final_grade.append([nota_final])

    school_absences_update = (
            sheet.values()
            .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="H4:H27", valueInputOption="RAW", body={"values": final_grade})
            .execute()
    )
    
    for value in abscenses:
        value_absences = int(value[0])
        grade = float(final_grade[abscenses.index(value)][0] * 1)
        #print(grade)
        if value_absences > 15:
            value[0] = "Reprovado por Falta"
        elif grade < 5 :
            value[0] = "Reprovado por Nota"
        elif grade >= 5 and grade < 7:
            value[0] = "Exame Final"
        elif grade >= 7:
            value[0] = "Aprovado"
        else :
            value[0] = ""

    school_absences_15_update = (
        sheet.values()
        .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="G4:G27", valueInputOption="RAW", body={"values": abscenses})
        .execute()
    )


  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()