from  googleapiclient.discovery import build
import xlsxwriter
import json

def excelHelper(ans):
    
    no_of_rows=len(ans)
    workbook = xlsxwriter.Workbook('output.xlsx')   
    worksheet = workbook.add_worksheet('Titles')
    worksheet.write(0, 0, 'Sr No')
    worksheet.write(0, 1, 'Title')
    

    worksheet.write(0,2, 'Status')
    
    for index, entry in enumerate(ans):
        print(entry)
        worksheet.write(index+1, 0, str(index+1))
        worksheet.write(index+1,1,entry)
        worksheet.data_validation(
        "C{}".format(index+2), {"validate": "list", "source": ["Completed", "In-Progress", "Not Important"]}
        )
    green_format=workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100'})
    red_format = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})
    worksheet.conditional_format("C{}:C{}".format(2,no_of_rows),{'type':     'cell',
                                    'criteria': 'equal to',
                                    'value':    '"Completed"',
                                    'format':   green_format})

    worksheet.conditional_format("C{}:C{}".format(2,no_of_rows),{'type':     'cell',
                                    'criteria': 'equal to',
                                    'value':    '"Not Important"',
                                    'format':   red_format})
    workbook.close()


API_KEY='YOUR_API_KEY'
youtube=build('youtube','v3',developerKey=API_KEY)

#Add the playlist that you want to search for
playlist_id='PLgUwDviBIf0q8Hkd7bK2Bpryj2xVJk8Vk'

request=youtube.playlistItems().list(
    part="snippet,contentDetails",
    playlistId=playlist_id,
    maxResults=200
)

response=request.execute()

with open("response.txt", "a") as fp:
    json.dump(response, fp)
    
ans=[]
n=len(response["items"])
for i in range(n):
    ans.append(response["items"][i]['snippet']['title'])

key='nextPageToken'
while key in response:
    next_page_token=response["nextPageToken"]
    request1=youtube.playlistItems().list(
        part="snippet,contentDetails",
        playlistId=playlist_id,
        maxResults=200,
        pageToken=next_page_token
    )
    response=request1.execute()
    print(response)
    with open("response.txt", "a") as fp:
        json.dump(response, fp)
    response["items"]=response.get("items","null") + response["items"]
 



# next_page_token=response["nextPageToken"]
# print(next_page_token)
# request1=youtube.playlistItems().list(
#     part="snippet,contentDetails",
#     playlistId='PLgUwDviBIf0q8Hkd7bK2Bpryj2xVJk8Vk',
#     maxResults=200,
#     pageToken=next_page_token
# )


# response1=request1.execute()
# print(response1["nextPageToken"])

# with open("response1.txt", "w") as fp:
#     json.dump(response1, fp)

n=len(response["items"])
for i in range(n):
    ans.append(response["items"][i]['snippet']['title'])

# n=len(response1["items"])
# for i in range(n):
#     ans.append(response1["items"][i]['snippet']['title'])

excelHelper(ans)


