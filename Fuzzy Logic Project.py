import pandas as pd
import warnings

warnings.filterwarnings("ignore")

def read_excel(path, sheet_target):
    data = pd.read_excel(path, sheet_name=sheet_target)
    return data.to_dict('record')

def output_xlsx(data_final, filename):
    df = pd.DataFrame(data_final, columns=["ID", "Service", "Price" ,"Defuzz"])
    try:
        df.to_excel(f'{filename}.xlsx')
        print(f"File saved as : {filename}.xlsx")
    except IOError:
        print("Could not open file! Please close Excel!")

def fuzz(n, data, fuzz_setting):
  choose = fuzz_setting['Method']
  final_data = []
  for i in range(0, n):
    low_fuzz = 0
    avg_fuzz = 0
    high_fuzz = 0
    # low
    if(data[i][choose] > fuzz_setting['low_top'] and data[i][choose] <= fuzz_setting['low_bot']):
      low_fuzz = abs(data[i][choose] - fuzz_setting['low_bot'])/abs(fuzz_setting['low_top']-fuzz_setting['low_bot'])
    elif(data[i][choose] <= fuzz_setting['low_top']) : low_fuzz = 1
    # avg
    if(data[i][choose] > fuzz_setting['avg_bot_left'] and data[i][choose] <= fuzz_setting['avg_top_left']):
      avg_fuzz = abs(data[i][choose] - fuzz_setting['avg_bot_left'])/abs(fuzz_setting['avg_bot_left']-fuzz_setting['avg_top_left'])
    elif(data[i][choose] > fuzz_setting['avg_top_left'] and data[i][choose] < fuzz_setting['avg_top_right']) : avg_fuzz = 1
    elif(data[i][choose] > fuzz_setting['avg_top_right'] and data[i][choose] <= fuzz_setting['avg_bot_right']) :
      avg_fuzz = abs(data[i][choose] - fuzz_setting['avg_bot_right'])/abs(fuzz_setting['avg_bot_right']-fuzz_setting['avg_top_right'])
    # low
    if(data[i][choose] > fuzz_setting['high_bot'] and data[i][choose] <= fuzz_setting['high_top']):
      high_fuzz = abs(data[i][choose] - fuzz_setting['high_bot'])/abs(fuzz_setting['high_top']-fuzz_setting['high_bot'])
    elif(data[i][choose] > fuzz_setting['high_top']) : high_fuzz = 1
    final = {'ID': data[i]['ID'], 'Low': low_fuzz, 'Average': avg_fuzz, 'High': high_fuzz}
    final_data.append(final)
  return final_data

def inference(n, fuzz_service, fuzz_price, inference_setting):
    inference_array = []
    for i in range(0,n):
        reject = []
        consider = []
        accept = []
        for j in inference_setting:
            if(j['Status'] == "Rejected"):
                take_minimum = min(fuzz_service[i][j['Service']], fuzz_price[i][j['Price']])
                reject.append(take_minimum)
            elif(j['Status'] == "Considered"):
                take_minimum = min(fuzz_service[i][j['Service']], fuzz_price[i][j['Price']])
                consider.append(take_minimum)
            elif(j['Status'] == "Accepted"):
                take_minimum = min(fuzz_service[i][j['Service']], fuzz_price[i][j['Price']])
                accept.append(take_minimum)
        result = {'ID': fuzz_service[i]['ID'], 'Rejected': max(reject), 'Considered': max(consider), 'Accepted': max(accept)}
        inference_array.append(result)
    return inference_array

def defuzz(sugeno, inference_data):
  defuzz_arr = []
  for i in inference_data:
    y = (i['Rejected'] * sugeno[0]) + (i['Considered']*sugeno[1]) + (i['Accepted']*sugeno[2]) / (i['Rejected'] + i['Considered'] + i['Accepted'] + 0.00000001)
    final = {'ID': i['ID'], 'Defuzz': y}
    defuzz_arr.append(final)
  return defuzz_arr

def bestof10(defuzz_data):
  final = sorted(defuzz_data, key=lambda i: i['Defuzz'], reverse=True)[0:10]
  for i in range(0,10):
    a = final[i]['ID']
    final[i] = {'ID': final[i]['ID'], 'Service': read_data[a-1]['Service'], 'Price': read_data[a-1]['Price'] ,'Defuzz': final[i]['Defuzz']}
  return final

def savefile(final_data):
    name = input("Choose filename : ")
    sort = input("Sort by ID or Defuzz? ")
    final_data = sorted(defuzz_data, key=lambda i: i['Defuzz'])
    if(sort == "ID"): final_data = sorted(defuzz_data, key=lambda i: i['ID'])
    elif(sort != "Defuzz") : print("Sorted by Defuzz by default")
    output_xlsx(finale, name)

inference_setting = [ 
{'Service': 'Low', 'Price': 'Low', 'Status': 'Rejected'}, {'Service': 'Low', 'Price': 'Average', 'Status': 'Rejected'}, {'Service': 'Low', 'Price': 'High', 'Status': 'Rejected'},
{'Service': 'Average', 'Price': 'Low', 'Status': 'Considered'}, {'Service': 'Average', 'Price': 'Average', 'Status': 'Considered'}, {'Service': 'Average', 'Price': 'High', 'Status': 'Rejected'},
{'Service': 'High', 'Price': 'Low', 'Status': 'Accepted'}, {'Service': 'High', 'Price': 'Average', 'Status': 'Accepted'}, {'Service': 'High', 'Price': 'High', 'Status': 'Considered'}]

service = {'low_top': 20, 'low_bot': 45, 'avg_top_left' : 55, 'avg_bot_left': 30, 
           'avg_top_right': 70, 'avg_bot_right': 75, 'high_top': 80, 'high_bot': 60, 'Method': 'Service'}

price = {'low_top': 2.0, 'low_bot': 4.0, 'avg_top_left' : 5, 'avg_bot_left': 3, 
           'avg_top_right': 7.0, 'avg_bot_right': 8.0, 'high_top': 9.0, 'high_bot': 6, 'Method': 'Price'}

sugeno = [50, 65, 80]

if __name__ == "__main__":
    read_data = read_excel('bengkel.xlsx', 'Sheet1')

    length = len(read_data)

    fuzz_service = fuzz(length, read_data, service)
    fuzz_price = fuzz(length, read_data, price)

    inference_data = inference(length, fuzz_service, fuzz_price, inference_setting)

    defuzz_data = defuzz(sugeno, inference_data)

    finale = bestof10(defuzz_data)
    finale = sorted(finale, key=lambda i: i['ID'])
    print("\n-----The Result-----")
    for i in finale:
        i = {'ID': i['ID'], 'Service': i['Service'], 'Price': i['Price'],'Defuzz': i['Defuzz']}
        print(i)

    savefile(finale)