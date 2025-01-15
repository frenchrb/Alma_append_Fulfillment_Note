import configparser
import json
import requests
import sys
import time
import xlrd
import xlwt
import xlutils.copy
from threading import Thread
from queue import Queue
from time import sleep
from ratelimit import limits, sleep_and_retry

# Read config file
config = configparser.ConfigParser()
config.read('local_settings.ini')
key = config['Alma Bibs R/W']['key']

num_worker_threads = 15
work_queue = Queue()
output_queue = Queue()

itemID_col_index = 1
holID_col_index = 2
bibID_col_index = 3
getitem_col_index = 4
fulfillment_note_col_index = 5
updateitem_col_index = 6

@sleep_and_retry
@limits(calls=15, period=1)
def api_request(type, item, json=None):
    if type == 'get':
        headers = {'accept':'application/json'}
        response = requests.get('https://api-na.hosted.exlibrisgroup.com/almaws/v1/bibs/'+item['bibID']+'/holdings/'+item['holID']+'/items/'+item['itemID']+'?view=brief&apikey='+key, headers=headers)
        return response
    if type == 'put':
        headers = {'accept':'application/json', 'Content-Type':'application/json'}
        response = requests.put('https://api-na.hosted.exlibrisgroup.com/almaws/v1/bibs/'+item['bibID']+'/holdings/'+item['holID']+'/items/'+item['itemID']+'?apikey='+key, headers=headers, data=json)
        return response

def worker():
    while True:
        item = work_queue.get()
        output = []
        get_response = api_request('get', item)
        output.append((item['row'], getitem_col_index, get_response.status_code))
        if get_response.status_code == 200:
            item_json = get_response.json()
            
            if item_json['item_data']['fulfillment_note'] != '':
                output.append((item['row'], fulfillment_note_col_index, item_json['item_data']['fulfillment_note']))
                new_note = item_json['item_data']['fulfillment_note'] + 'INSERT TEXT TO ADD TO FULFILLMENT NOTE HERE'
            else:
                new_note = 'INSERT TEXT TO ADD TO FULFILLMENT NOTE HERE'
            
            item_json['item_data']['fulfillment_note'] = new_note
            put_response = api_request('put', item, json.dumps(item_json))
            output.append((item['row'], updateitem_col_index, put_response.status_code))
            
        output_queue.put(output)
        print(item['row'], item['itemID'])
        work_queue.task_done()

def out_worker(book_in, input):
    # Copy spreadsheet for output
    book_out = xlutils.copy.copy(book_in)
    sheet_out = book_out.get_sheet(0)
    
    # Add new column headers
    sheet_out.write(0,fulfillment_note_col_index,'Existing_Fulfillment_Note')
    sheet_out.write(0,getitem_col_index,'Get_Item')
    sheet_out.write(0,updateitem_col_index,'Update_Item')
    
    while True:
        items = output_queue.get()
        for item in items:
            sheet_out.write(item[0], item[1], item[2])
        book_out.save(input+'_results.xls')
        output_queue.task_done()

def main(input):
    st = time.localtime()
    start_time = time.strftime("%H:%M:%S", st)
   
    # Read spreadsheet
    book_in = xlrd.open_workbook(input)
    sheet1 = book_in.sheet_by_index(0) #get first sheet

    Thread(target=out_worker, args=(book_in, input,), daemon=True).start()
    for i in range(num_worker_threads):
        Thread(target=worker, daemon=True).start()
    

    for row in range(1, sheet1.nrows):
        item = {}
        item['row'] = row
        item['bibID'] = sheet1.cell(row, bibID_col_index).value
        item['holID'] = sheet1.cell(row, holID_col_index).value
        item['itemID'] = sheet1.cell(row, itemID_col_index).value
        work_queue.put(item)

    work_queue.join()
    output_queue.join()

    et = time.localtime()
    end_time = time.strftime("%H:%M:%S", et)
    print('Start Time: ', start_time)
    print('End Time: ', end_time)   

if __name__ == '__main__':
    main(sys.argv[1])