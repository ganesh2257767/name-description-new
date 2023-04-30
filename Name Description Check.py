import PySimpleGUI as sg
import requests
import json
import pandas as pd
import os

# TODO: 
# Sanitize input

icon_path = os.getcwd()

uow_uat = 'https://ws-uat.suddenlink.com/optimum-online-order-ws/rest/OfferService/getBundles'
uow_uat1 = 'https://ws-uat.suddenlink.com/uat1/optimum-online-order-ws/rest/OfferService/getBundles'
uow_uat2 = 'https://ws-uat.suddenlink.com/uat2/optimum-online-order-ws/rest/OfferService/getBundles'
dsa_uat = 'https://ws-uat.suddenlink.cequel3.com/optimum-ecomm-abstraction-ws/rest/uow/searchProductOffering'
dsa_uat1 = 'https://ws-uat.suddenlink.cequel3.com/uat1/optimum-ecomm-abstraction-ws/rest/uow/searchProductOffering'
dsa_uat2 = 'https://ws-uat.suddenlink.cequel3.com/uat2/optimum-ecomm-abstraction-ws/rest/uow/searchProductOffering'

def check(url, req, chann, data, ch_param, area):
    try:
        if chann == 'uow':
            res = requests.post(url, json=req, auth=('unittest', 'test01')).json()
            offers = res["productOfferings"]["productOfferingResults"]
            
        elif chann == 'dsa':
            res = requests.post(url, json=req, verify=False).json()
            offers = res["searchProductOfferingReturn"]["productOfferingResults"]
        if not offers:
            raise KeyError
    except KeyError:
        sg.Popup("No offers returned!\n\nCheck input values and try again.")
    except requests.exceptions.RequestException as e:
        print(e)
        sg.Popup("Failed to establish a new connection.\n\nCheck VPN.")
    else:
        offer_name_desc = {}
        
        if values['mobile_checkbox']:
            if len(offers) == 1:
                if isinstance(offers, list):
                    offer_id = str(offers[0]["matchingProductOffering"]["ID"])
                    offer_name = offers[0]["matchingProductOffering"]["mobileTitle"]
                    offer_desc = offers[0]["matchingProductOffering"]["mobileDescription"]
                    offer_price = str(offers[0]["matchingProductOffering"].get("mobileStartingPrice", ""))
                    try:
                        offer_price = f'{float(offer_price):.2f}'
                    except:
                        pass
                    offer_name_desc[offer_id] = [offer_name, offer_desc, offer_price]
                else:
                    offer_id = str(offers["matchingProductOffering"]["ID"])
                    offer_name = offers["matchingProductOffering"]["mobileTitle"]
                    offer_desc = offers["matchingProductOffering"]["mobileDescription"]
                    offer_price = str(offers["matchingProductOffering"].get("mobileStartingPrice", ""))
                    try:
                        offer_price = f'{float(offer_price):.2f}'
                    except:
                        pass
                    offer_name_desc[offer_id] = [offer_name, offer_desc, offer_price]
            else:
                for offer in offers:
                    print(offer)
                    offer_id = str(offer["matchingProductOffering"]["ID"])
                    offer_name = offer["matchingProductOffering"]["mobileTitle"]
                    offer_desc = offer["matchingProductOffering"]["mobileDescription"]
                    offer_price = str(offer["matchingProductOffering"].get("mobileStartingPrice", ""))
                    try:
                        offer_price = f'{float(offer_price):.2f}'
                    except:
                        pass
                    offer_name_desc[offer_id] = [offer_name, offer_desc, offer_price]
        else:
            if len(offers) == 1:
                if isinstance(offers, list):
                    offer_id = str(offers[0]["matchingProductOffering"]["ID"])
                    offer_name = offers[0]["matchingProductOffering"]["title"]
                    offer_desc = offers[0]["matchingProductOffering"]["description"]
                    offer_price = str(offers[0]["matchingProductOffering"]["startingPrice"])
                    offer_name_desc[offer_id] = [offer_name, offer_desc, f'{float(offer_price):.2f}']
                else:
                    offer_id = str(offers["matchingProductOffering"]["ID"])
                    offer_name = offers["matchingProductOffering"]["title"]
                    offer_desc = offers["matchingProductOffering"]["description"]
                    offer_price = str(offers["matchingProductOffering"]["startingPrice"])
                    offer_name_desc[offer_id] = [offer_name, offer_desc, f'{float(offer_price):.2f}']
            else:
                for offer in offers:
                    offer_id = str(offer["matchingProductOffering"]["ID"])
                    offer_name = offer["matchingProductOffering"]["title"]
                    offer_desc = offer["matchingProductOffering"]["description"]
                    offer_price = str(offer["matchingProductOffering"]["startingPrice"])
                    offer_name_desc[offer_id] = [offer_name, offer_desc, f'{float(offer_price):.2f}']
            
        actual_name = []
        actual_description = []
        actual_price = []
        result = []
        for oid, name, descr, price in zip(data['Offer ID'], data['Offer Name'], data['Offer Description'], data['Offer Price']):
            if oid in offer_name_desc.keys():
                actual_name.append(offer_name_desc[oid][0])
                actual_description.append(offer_name_desc[oid][1])
                actual_price.append(offer_name_desc[oid][2])
                if ch_param == 'all':                    
                    if descr.strip() == offer_name_desc[oid][1] and name.strip() == offer_name_desc[oid][0] and f'{float(price):.2f}' == f'{float(offer_name_desc[oid][2]):.2f}':
                        result.append('Pass')
                    else:
                        result.append('Fail')
                elif ch_param == 'ch_desc':
                    if str(descr).strip() == offer_name_desc[oid][1] and name.strip() == offer_name_desc[oid][0]:
                        result.append('Pass')
                    else:
                        result.append('Fail')
                elif ch_param == 'ch_price':
                    
                    if name.strip() == offer_name_desc[oid][0] and f'{float(price):.2f}' == f'{float(offer_name_desc[oid][2]):.2f}':
                        result.append('Pass')
                    else:
                        result.append('Fail')
                else:
                    if name.strip() == offer_name_desc[oid][0]:
                        result.append('Pass')
                    else:
                        result.append('Fail')
            else:
                actual_name.append("Not found")
                actual_description.append("Not found")
                actual_price.append("Not found")
                result.append('NA')
        
        data["Actual Name"] = actual_name
        data["Actual Description"] = actual_description
        data["Actual Price"] = actual_price
        data["Result"] = result
        
        
        pass_data = data[data.Result == 'Pass']
        fail_data = data[data.Result == 'Fail']
        new_input = data[data.Result == 'NA']
        area = '_'.join(area)
        
        try:
            pass_data.to_csv(out_path + f'/{area}_pass.csv', mode='a', index=False)
            fail_data.to_csv(out_path + f'/{area}_fail.csv', mode='a', index=False)
            new_input.to_excel(out_path + f'/{area}_NA.xlsx', index=False)
        except PermissionError:
            sg.Popup("File already open or present in a folder without permissions!")
        else:
            sg.Popup(f'Statistics of run\nTotal runs: {len(result)}\nPassed: {result.count("Pass")}\nFailed: {result.count("Fail")}\nNA: {result.count("NA")}')
            window['-OPEN-'].update(disabled=False)

flag = 'true'
opt_markets = ['K', 'M', 'N', 'G']
sdl_markets = ['A', 'B', 'C', 'E', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'V']
cluster_sdl = [10, 21, 58, 59, 66, 67, 90, 91, 92, 93, 95]
cluster_opt = [6, 10, 86]

layout0 = [
    [sg.Text('Browse or enter the path to input file.'), sg.Text('Hover here for Help', relief='raised', tooltip='''Create an excel file with 4 columns.\n
All the columns should have an heading, the heading names don't matter.\n
The first column should have the offer IDs.\n
The second column must have the offer names. (Note: SDL sheet has different columns for ISA/DSA names and UOW names.)\n
The third column must have offer descriptions. (Like offer names, there are different columns for ISA/DSA and UOW descriptions).\n
The fourth column should have the prices. Now depending on what cluster you will use, use the correct prices (Comp/Non Comp)\n
You don't need to specify complete address, just provide Corp, Market, Cluster, Ftax and EID. (Use in parallel with the EID tool for efficiency.)\n
Reach out to me for further assistance.''')],
    [sg.In(key='input_file'), sg.FileBrowse(key='file', target='input_file')],
    [sg.Button('Upload', key='upload')]
]

layout = [
    [sg.Text("Select Area", size = (15, 1)), sg.Radio("Optimum", group_id="area", key="opt", enable_events=True, default=True, size=(6, 1)), sg.Radio("Suddenlink", group_id="area", key="sdl", enable_events=True)],
    [sg.Text("Select Channel", size = (15, 1)), sg.Radio("ISA/DSA", group_id="channel", key="dsa", enable_events=True, default=True, size=(6, 1)), sg.Radio("UOW", group_id="channel", key="uow", enable_events=True)],
    [sg.Text("Select Environment", size = (15, 1)), sg.Radio("UAT", group_id="env", key="uat", enable_events=True, default=True, size=(6, 1)), sg.Radio("UAT1", group_id="env", key="uat1", enable_events=True, default=True, size=(6, 1)), sg.Radio("UAT2", group_id="env", key="uat2", enable_events=True)],
    [sg.Text("Promotional?", size = (15, 1)), sg.Radio("Promo", group_id="promo", key="pro", enable_events=True, default=True, size=(6, 1)), sg.Radio("Full Rate", group_id="promo", key="full", enable_events=True)],
    [sg.Text("Enter Corp", size = (15, 1)), sg.InputText(key='corp', size=(8, 1))],
    [sg.Text("Select Market", size = (15, 1)), sg.DropDown(values = opt_markets, key='market', size=(6, 1))],
    [sg.Text("Select Cluster", size = (15, 1)), sg.DropDown(values = cluster_opt, key='cluster', size=(6, 1))],
    [sg.Text("Enter FTAX", size = (15, 1)), sg.InputText(key='ftax', visible=False, size=(8, 1))],
    [sg.Text("Enter EID", size = (15, 1)), sg.InputText(key='eid', visible=False, size=(8, 1))],
    [sg.Text("Check Parameters", size = (15, 1)),
     sg.Checkbox('Offer ID', default=True, disabled=True, size=(8, 1), background_color='white'),
     sg.Checkbox('Name', default=True, disabled=True,
                 size=(8, 1), background_color='white')],
    [sg.Text("", size=(15, 1)), sg.Checkbox('Description', enable_events=True, key='check_description', size=(8, 1)),
     sg.Checkbox('Price', enable_events=True, key='check_price', size=(8, 1))],
    [sg.Text("Mobile offers?", size=(15, 1)), sg.Checkbox('', key='mobile_checkbox')],
    [sg.Button('Upload another file', key='-ANOTHER-'), sg.Submit("Check", key='-SUBMIT-',  size=(10, 1)), sg.Button('Open output folder', key='-OPEN-', disabled=True)]
]

main_layout = [
    [sg.Column(layout0, key='-COL1-'),
    sg.Column(layout, visible=False, key='-COL2-')]
]

window = sg.Window('Offer Name/Description Checker', main_layout, icon=icon_path+'\\main_icon.ico')


while True:
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == 'Cancel':
        break

    if event == 'upload':
        if values['file']:
            path = values['file']
            out_path = '/'.join(values['file'].split('/')[:-1])
            try:
                data = pd.read_excel(path, usecols = [0, 1, 2, 3], names = ["Offer ID", "Offer Name", "Offer Description", "Offer Price"], dtype={"Offer ID": str, "Offer Name": str, "Offer Description": str, "Offer Price": float})
            except ValueError as e:
                if e == "Number of passed names did not match number of header fields in the file":
                    sg.Popup('Input file not in correct format.')
                elif e == "Excel file format cannot be determined, you must specify an engine manually.":
                    sg.Popup('Please select an excel file.')
                    
            else:
                window['-COL1-'].update(visible=False)
                window['-COL2-'].update(visible=True)
        else:
            sg.Popup('Please select an input file!')
    
    if event == 'sdl':
        window['market'].update(values=sdl_markets)
        window['cluster'].update(values=cluster_sdl)
        window['ftax'].update(visible=True)
        window['eid'].update(visible=True)

    if event == 'opt':
        window['market'].update(values=opt_markets)
        window['cluster'].update(values=cluster_opt)
        window['ftax'].update(visible=False)
        window['eid'].update(visible=False)

    if event == 'uow':
        window['pro'].update(disabled=True)
        window['full'].update(disabled=True)
    
    if event == 'dsa':
        window['pro'].update(disabled=False)
        window['full'].update(disabled=False)
        
    if event == 'pro':
        flag = 'true'
        
    if event == 'full':
        flag = 'false'
        
    if event == '-SUBMIT-':
        if values['sdl'] and values['uow']:
            if all([values['market'], values['corp'], values['cluster'], values['ftax'], values['eid']]):
                area = (values['corp'], values['ftax'], values['eid'], values['market'])
                if values['uat']:
                    url = uow_uat
                elif values['uat1']:
                    url = uow_uat1
                else:
                    url = uow_uat2
                payload = f'''{{"productOfferingsRequest":{{"customerInteractionId":"1228012","eligibilityID": "{values['eid']}","accountDetails":{{"clust":"{values['cluster']}","corp":"{values['corp']}","cust":"1","eligibilityId": "test","ftax":"{values['ftax']}","hfstatus":"3","house":"test","id":0,"mkt":"{values['market']}","service_housenbr":"test","servicestreetaddr":"test","service_aptn": "test","service_city":"test","service_state":"test","service_zipcode":"test","tdrop": "O"}},"newCustomer":true,"sessionId":"LDPDPJCBBH08VVL9KKY","shoppingCartId":"FTJXQYDN","footprint": "suddenlink"}}}}'''
                req = json.loads(payload)
                if values['check_description'] and values['check_price']:
                    ch_param = 'all'
                elif values['check_description']:
                    ch_param = 'ch_desc'
                elif values['check_price']:
                    ch_param = 'ch_price'
                else:
                    ch_param = ''
                check(url, req, 'uow', data, ch_param, area)
            else:
                sg.Popup('Enter all the values!')

        elif values['opt'] and values['uow']:
            if all([values['market'], values['corp'], values['cluster']]):
                area = (values['corp'], values['ftax'], values['eid'], values['market'])
                if values['uat']:
                    url = uow_uat
                elif values['uat1']:
                    url = uow_uat1
                else:
                    url = uow_uat2
                if values['market'] in ('G', 'F'):
                    payload = f'''{{"productOfferingsRequest":{{"customerInteractionId":"1228012","accountDetails":{{"clust":"{values['cluster']}","corp":"{values['corp']}","cust":"1","disconnectedDate": "2018-08-31T00:00:00-05:00","ftax":"72","hfstatus":"3","house":"test","id":0,"mkt":"{values['market']}","service_housenbr":"53","service_apt": "APT","servicestreetaddr":"53 SLEEPY LN","service_aptn": "55","service_city":"HICKSVILLE","service_state":"NY","service_zipcode":"11801"}},"newCustomer":true,"sessionId":"LDPDPJCBBH08VVL9KKY","shoppingCartId":"FTJXQYDN"}}}}'''
                else:
                    payload = f'''{{"productOfferingsRequest":{{"customerInteractionId":"1228012","accountDetails":{{"clust":"{values['cluster']}","corp":"{values['corp']}","cust":"1","disconnectedDate": "2018-08-31T00:00:00-05:00","ftax":"72","hfstatus":"3","house":"test","id":0,"mkt":"{values['market']}","service_housenbr":"test","service_apt": "test","servicestreetaddr":"test","service_aptn": "test","service_city":"test","service_state":"test","service_zipcode":"test"}},"newCustomer":true,"sessionId":"LDPDPJCBBH08VVL9KKY","shoppingCartId":"FTJXQYDN"}}}}'''
    
                req = json.loads(payload)
                if values['check_description'] and values['check_price']:
                    ch_param = 'all'
                elif values['check_description']:
                    ch_param = 'ch_desc'
                elif values['check_price']:
                    ch_param = 'ch_price'
                else:
                    ch_param = ''
                check(url, req, 'uow', data, ch_param, area)
            else:
                sg.Popup('Enter all the values!')
        
        elif values['sdl'] and values['dsa']:
            if all([values['market'], values['corp'], values['cluster'], values['ftax'], values['eid']]):
                area = (values['corp'], values['ftax'], values['eid'], values['market'])
                if values['uat']:
                    url = dsa_uat
                elif values['uat1']:
                    url = dsa_uat1
                else:
                    url = dsa_uat2
                payload = f'''{{"salesContext":{{"localeString":"en_US","salesChannel":"DSL"}},"searchProductOfferingFilterInfo":{{"oolAvailable":true,"ovAvailable":true,"ioAvailable":true,"includeExpiredOfferings":false,"salesRuleContext":{{"customerProfile":{{"anonymous":true}},"customerInfo":{{"customerType":"R","newCustomer":true,"orderType":"Install","isPromotion":{flag},"eligibilityID":"{values['eid']}"}}}},"eligibilityStatus":[{{"code":"EA"}}]}},"offeringReadMask":{{"value":"SUMMARY"}},"checkCustomerProductOffering":false,"locale":"en_US","cartId":"FTJXQYDN","serviceAddress":{{"apt":"test","fta":"{values['ftax']}","street":"test","city":"test","state":"test","zipcode":"test","type":"","clusterCode":"{values['cluster']}","mkt":"{values['market']}","corp":"{values['corp']}","house":"test","cust":"1"}},"generics":false}}'''
                req = json.loads(payload)
                if values['check_description'] and values['check_price']:
                    ch_param = 'all'
                elif values['check_description']:
                    ch_param = 'ch_desc'
                elif values['check_price']:
                    ch_param = 'ch_price'
                else:
                    ch_param = ''
                check(url, req, 'dsa', data, ch_param, area)
            else:
                sg.Popup('Enter all the values!')

        elif values['opt'] and values['dsa']:
            if all([values['market'], values['corp'], values['cluster']]):
                area = (values['corp'], values['ftax'], values['eid'], values['market'])
                if values['uat']:
                    url = dsa_uat
                elif values['uat1']:
                    url = dsa_uat1
                else:
                    url = dsa_uat2
                payload = f'''{{"salesContext":{{"localeString":"en_US","salesChannel":"DSL"}},"searchProductOfferingFilterInfo":{{"oolAvailable":true,"ovAvailable":true,"ioAvailable":true,"includeExpiredOfferings":false,"salesRuleContext":{{"customerProfile":{{"anonymous":true}},"customerInfo":{{"customerType":"R","newCustomer":true,"orderType":"Install","isPromotion":{flag},"eligibilityID":"test"}}}},"eligibilityStatus":[{{"code":"EA"}}]}},"offeringReadMask":{{"value":"SUMMARY"}},"checkCustomerProductOffering":false,"locale":"en_US","cartId":"FTJXQYDN","serviceAddress":{{"apt":"test","fta":"40","street":"test","city":"test","state":"test","zipcode":"test","type":"","clusterCode":"{values['cluster']}","mkt":"{values['market']}","corp":"{values['corp']}","house":"test","cust":"1"}},"generics":false}}'''
                req = json.loads(payload)
                if values['check_description'] and values['check_price']:
                    ch_param = 'all'
                elif values['check_description']:
                    ch_param = 'ch_desc'
                elif values['check_price']:
                    ch_param = 'ch_price'
                else:
                    ch_param = ''
                check(url, req, 'dsa', data, ch_param, area)
            else:
                sg.Popup('Enter all the values!')


    if event == '-OPEN-':
        os.startfile(out_path)

    if event == '-ANOTHER-':
        window['-COL2-'].update(visible=False)
        window['-COL1-'].update(visible=True)
        

window.close()
