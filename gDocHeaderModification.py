from __future__ import print_function
import pickle
import os.path
import csv

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# WARNING IF USED INCORRECTLY THIS CAN DELETE ALL THE HEADERS IN A FOLDER
# For adding set to True for Removing set to False
AddingHeader = True

message = "Looking for COVID support and guidance? Visit our COVID Contingency Materials Summary Page.\n"
urlLink = 'https://elearningyork.wordpress.com/2020/03/13/transferring-face-to-face-teaching-to-online/'
messageLength = len(message)

def changeHeader(item):
    headerLocation = docService.documents().get(documentId=item['id'],fields ='headers').execute()
    headerId = list(headerLocation.get('headers'))[0]
    contentLength = len(list(list(headerLocation['headers'][headerId]['content'])))
    finalSection = len(list(list(headerLocation['headers'][headerId]['content'])[contentLength - 1]['paragraph']['elements'])) - 1
    endIndex = list(list(headerLocation['headers'][headerId]['content'])[contentLength - 1]['paragraph']['elements'])[finalSection]['endIndex'] - 1

    textContent = []
    for contentInt in range(contentLength):
        for section in range(finalSection + 1):
            try:
                sectionText = list(headerLocation['headers'][headerId]['content'][contentInt]['paragraph']['elements'])[section]['textRun']['content']
                textContent.append(sectionText)
            except:
                True
    headerText = "".join(textContent)
    print("Header Text: " + headerText)

    messageInHeader = False

    if(headerText.find(message) != -1):
        print("Message Found!")
        print(headerText.find(message))
        messageInHeader = True


    if (AddingHeader and messageInHeader != True):
        requests = [
            {
                'insertText': {
                    'location': {
                        'segmentId': headerId,
                        'index': endIndex,
                    },
                    'text': message
                }
            },
            {
                'updateTextStyle': {
                    'range': {
                        'segmentId': headerId,
                        'startIndex': endIndex,
                        'endIndex': 40 + endIndex
                    },
                'textStyle': {
                    'foregroundColor': {
                        'color': {
                            'rgbColor': {
                                'blue': 0.0,
                                'green': 0.0,
                                'red': 1.0
                            }
                        }
                    },
                    'bold': True,
                },
                    'fields': 'foregroundColor,bold'
                }
            },
        {
            'updateTextStyle': {
                'range': {
                    'segmentId': headerId,
                    'startIndex': 47 + 1,
                    'endIndex': messageLength + endIndex
                },
                'textStyle': {
                    'link': {
                        'url': urlLink
                    }
                },
                'fields': 'link'
            }
        }]
        result = docService.documents().batchUpdate(documentId=item['id'], body={'requests': requests}).execute()
    if(AddingHeader == False and messageInHeader):
        requests = [
        {
            'deleteContentRange': {
                'range': {
                    'segmentId': headerId,
                    'startIndex': endIndex - messageLength,
                    'endIndex': endIndex,
                }

            }

        },
    ]
        result = docService.documents().batchUpdate(documentId=item['id'], body={'requests': requests}).execute()

def updateTopOfPage(item):
    fileLocation = docService.documents().get(documentId=item['id']).execute()
    endIndex = 1


    # textContent = []
    # for contentInt in range(contentLength):
    #     for section in range(finalSection + 1):
    #         try:
    #             sectionText = \
    #             list(headerLocation['headers'][headerId]['content'][contentInt]['paragraph']['elements'])[section][
    #                 'textRun']['content']
    #             textContent.append(sectionText)
    #         except:
    #             True
    # headerText = "".join(textContent)
    # print("Header Text: " + headerText)



    # if (headerText.find(message) != -1):
    #     print("Message Found!")
    #     print(headerText.find(message))
    #     messageInBody = True

    if (AddingHeader):
        requests = [
            {
                'insertText': {
                    'location': {
                        'index': endIndex,
                    },
                    'text': message
                }
            },
            {
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': endIndex,
                        'endIndex': messageLength + endIndex
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'NORMAL_TEXT',
                        'alignment':'CENTER'
                    },
                    'fields': 'namedStyleType,alignment'
                },
            },
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': endIndex,
                        'endIndex': messageLength + endIndex
                    },
                    'textStyle': {
                        'fontSize': {
                              "magnitude": 9,
                              "unit": "PT"
                            }
                    },
                    'fields': 'fontSize'
                },
            },
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': endIndex,
                        'endIndex': 39 + endIndex
                    },
                    'textStyle': {
                        'foregroundColor': {
                            'color': {
                                'rgbColor': {
                                    'blue': 0.0,
                                    'green': 0.0,
                                    'red': 1.0
                                }
                            }
                        },
                        'bold': True,
                    },
                    'fields': 'foregroundColor,bold'
                }
            },
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': 46 + 1,
                        'endIndex': messageLength + endIndex
                    },
                    'textStyle': {
                        'link': {
                            'url': urlLink
                        }
                    },
                    'fields': 'link'
                }
            }]
        result = docService.documents().batchUpdate(documentId=item['id'], body={'requests': requests}).execute()
    if (AddingHeader == False):
        requests = [
            {
                'replaceAllText': {
                    'containsText': {
                        'text': message,
                        'matchCase': 'true'
                    },
                    'replaceText': "",
                }
            }
            # {
            #     'deleteContentRange': {
            #         'range': {
            #             'startIndex': endIndex ,
            #             'endIndex': endIndex + messageLength,
            #         }
            #
            #     }
            #
            # },
        ]
        result = docService.documents().batchUpdate(documentId=item['id'], body={'requests': requests}).execute()


def createLink(id, mimeType):
    switcher ={
        'application/vnd.google-apps.spreadsheet':'https://docs.google.com/spreadsheets/d/',
        'application/vnd.google-apps.document':'https://docs.google.com/document/d/',
        'application/vnd.google-apps.presentation':'https://docs.google.com/presentation/d/',
        'application/vnd.google-apps.folder':'https://drive.google.com/drive/folders/',
        'application/vnd.google-apps.file':'https://drive.google.com/file/d/'
    }
    linkStart = switcher.get(mimeType, 'https://drive.google.com/file/d/')
    link = linkStart + id
    return link

def fileType(mimeType):
    switcher ={
        'application/vnd.google-apps.spreadsheet':'sheet',
        'application/vnd.google-apps.document':'doc',
        'application/vnd.google-apps.folder':'folder',
        'application/vnd.google-apps.presentation': 'slides',
        'application/vnd.google-apps.file':'file'
    }
    file = switcher.get(mimeType, mimeType)
    return file



def writeCSV(csv_writer, items,folderName):
    for item in items:
        print(u'{0} ({1})'.format(item['name'], item['id']))
        pageTitle = '"' + item['name'] + '"'
        fileId = item['id']
        # id = item['id']
        mimeType = fileType(item['mimeType'])
        link = createLink(fileId, item['mimeType'])
        folderNameUpdate = '"' + folderName +'"'
        if(item['mimeType']!='application/vnd.google-apps.folder'):
            if(item['mimeType']=='application/vnd.google-apps.document'):
                try:
                    updateTopOfPage(item)
                except:
                    print('Error!')

            csv_writer.writerow(
                {'page_title': pageTitle, 'type': mimeType, 'link': link, 'folder_name': folderNameUpdate})
        if(item['mimeType']=='application/vnd.google-apps.folder'):
            newfolderName = folderName + "/" + item['name']
            formattedFolderId = "'{}'".format(item['id'])
            query = formattedFolderId + ' in parents'
            formattedQuery = '"{}"'.format(query)
            # Call the Drive v3 API
            newresults = service.files().list(q=query,
                                           pageSize=1000,
                                           fields="nextPageToken, files(kind,id,name,mimeType,parents)").execute()
            newitems = newresults.get('files', [])
            print(newfolderName)
            # print('Do you wish to process this folder?')
            # processFolder = input()
            processFolder = 'y'
            if processFolder.lower() == 'y':
                writeCSV(csv_writer,newitems,newfolderName)

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly','https://www.googleapis.com/auth/documents']

creds = None
csvLocation = 'links.csv'

# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('drive', 'v3', credentials=creds)
docService = build('docs', 'v1', credentials=creds)

folderAddress = input('Enter Folder Address: ')
folderName = "TEST"
print('THIS PROGRAM WILL CHANGE THE HEADERS AND MAY DELETE ALL HEADERS ARE YOU SURE?')
print('---------')
verify = input('ARE YOU SURE? CHECK SETTINGS BEFORE CONTINUING. PRESS Y to CONTINUE or HIT ANY KEY TO EXIT ')
if(verify.lower() != 'y'):
    exit()
folderId = folderAddress.split('/folders/')[1]
formattedFolderId = "'{}'".format(folderId)
query = formattedFolderId + ' in parents'
formattedQuery = '"{}"'.format(query)

# Call the Drive v3 API
results = service.files().list(q=query,
    pageSize=1000, fields="nextPageToken, files(kind,id,name,mimeType,parents)").execute()
items = results.get('files', [])

if not items:
    print('No files found.')
else:
    print('Files:')

    with open(csvLocation, mode='w') as csvFile:
        fieldnames = ['folder_name','page_title', 'type', 'link','id']
        csv_writer = csv.DictWriter(csvFile, fieldnames=fieldnames)
        csv_writer.writeheader()
        writeCSV(csv_writer,items,folderName)

        # Announce that we've finished
        print("Finished! Written to CSV.")