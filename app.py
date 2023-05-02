version = 'v1.1'

import gooeypie as gp
import pandas as pd
import json
import requests
import threading
import sys
import logging
import traceback
import os
import tkinter as tk

logger = logging.getLogger(__name__)
handler = logging.FileHandler('logs.txt')
handler.setLevel(logging.DEBUG)
format = logging.Formatter('[%(asctime)s] - [%(name)s] - [%(funcName)s:%(lineno)d] - [%(levelname)s] - %(message)s')
handler.setFormatter(format)
logger.addHandler(handler)

headers = ['Offer ID', "Gathering Name", "Gathering Description", "Gathering Price", "EPC Name", "EPC Description", "EPC Price", "Result"]

file_path: str = None
offers: dict = None
final_dict: dict = {}
data: dict = None


def check_version():
    response = requests.get('https://raw.githubusercontent.com/ganesh2257767/name-description-new/main/app.py')
    if response.status_code == 200:
        version = response.content.decode().split('\n')[0]
        print(version)
    else:
        print("Cannot check for updates now.")


def handle_thread_exception(args):
    if issubclass(args.exc_type, KeyboardInterrupt):
        sys.__excepthook__(args.exc_type, args.exc_value, args.exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
    app.alert("Exception", f'Uncaught exception:\nType: {args.exc_type}\nValue: {args.exc_value}\nTraceback: {traceback.format_tb(args.exc_traceback)}', "error")


def handle_exceptions(*args):
    if issubclass(args[1], KeyboardInterrupt):
        sys.__excepthook__(args[1], args[2], args[3])
        return

    logger.error("Uncaught exception", exc_info=(args[1], args[2], args[3]))
    app.alert("Exception", f'Uncaught exception:\nType: {args[1]}\nValue: {args[2]}\nTraceback: {traceback.format_tb(args[3])}', "error")


urls: dict = {
    'uow': {
        'uat': 'https://ws-uat.suddenlink.com/optimum-online-order-ws/rest/OfferService/getBundles',
        'uat1': 'https://ws-uat.suddenlink.com/uat1/optimum-online-order-ws/rest/OfferService/getBundles',
        'uat2': 'https://ws-uat.suddenlink.com/uat2/optimum-online-order-ws/rest/OfferService/getBundles'
    },
    'dsa': {
        'uat': 'https://ws-uat.suddenlink.cequel3.com/optimum-ecomm-abstraction-ws/rest/uow/searchProductOffering',
        'uat1': 'https://ws-uat.suddenlink.cequel3.com/uat1/optimum-ecomm-abstraction-ws/rest/uow/searchProductOffering',
        'uat2': 'https://ws-uat.suddenlink.cequel3.com/uat2/optimum-ecomm-abstraction-ws/rest/uow/searchProductOffering'
    }
}

payloads: dict = {
    'uow': {
        'opt': '''{{"productOfferingsRequest":{{"customerInteractionId":"1228012","accountDetails":{{"clust":"{}","corp":"{}","cust":"1","disconnectedDate": "2018-08-31T00:00:00-05:00","ftax":"72","hfstatus":"3","house":"test","id":0,"mkt":"{}","service_housenbr":"{}","service_apt": "test","servicestreetaddr":"{}","service_aptn": "test","service_city":"{}","service_state":"{}","service_zipcode":"{}"}},"newCustomer":true,"sessionId":"LDPDPJCBBH08VVL9KKY","shoppingCartId":"FTJXQYDN"}}}}''',
        'sdl': '''{{"productOfferingsRequest":{{"customerInteractionId":"1228012","eligibilityID": "{}","accountDetails":{{"clust":"{}","corp":"{}","cust":"1","ftax":"{}","hfstatus":"3","house":"test","id":0,"mkt":"{}","service_housenbr":"{}","servicestreetaddr":"{}","service_aptn": "test","service_city":"{}","service_state":"{}","service_zipcode":"{}","tdrop": "O"}},"newCustomer":true,"sessionId":"LDPDPJCBBH08VVL9KKY","shoppingCartId":"FTJXQYDN","footprint": "suddenlink"}}}}'''
    },
    'dsa': {
        'opt': '''{{"salesContext":{{"localeString":"en_US","salesChannel":"DSL"}},"searchProductOfferingFilterInfo":{{"oolAvailable":true,"ovAvailable":true,"ioAvailable":true,"includeExpiredOfferings":false,"salesRuleContext":{{"customerProfile":{{"anonymous":true}},"customerInfo":{{"customerType":"R","newCustomer":true,"orderType":"Install","isPromotion":{},"eligibilityID":"test"}}}},"eligibilityStatus":[{{"code":"EA"}}]}},"offeringReadMask":{{"value":"SUMMARY"}},"checkCustomerProductOffering":false,"locale":"en_US","cartId":"FTJXQYDN","serviceAddress":{{"apt":"test","fta":"40","street":"test","city":"test","state":"test","zipcode":"test","type":"","clusterCode":"{}","mkt":"{}","corp":"{}","house":"test","cust":"1"}},"generics":false}}''',
        'sdl': '''{{"salesContext":{{"localeString":"en_US","salesChannel":"DSL"}},"searchProductOfferingFilterInfo":{{"oolAvailable":true,"ovAvailable":true,"ioAvailable":true,"includeExpiredOfferings":false,"salesRuleContext":{{"customerProfile":{{"anonymous":true}},"customerInfo":{{"customerType":"R","newCustomer":true,"orderType":"Install","isPromotion":{},"eligibilityID":"{}"}}}},"eligibilityStatus":[{{"code":"EA"}}]}},"offeringReadMask":{{"value":"SUMMARY"}},"checkCustomerProductOffering":false,"locale":"en_US","cartId":"FTJXQYDN","serviceAddress":{{"apt":"test","fta":"{}","street":"test","city":"test","state":"test","zipcode":"test","type":"","clusterCode":"{}","mkt":"{}","corp":"{}","house":"test","cust":"1"}},"generics":false}}'''
    }
}

corps: dict = {
    ('7801', '7816'): '61 SLEEPY LN HICKSVILLE NY 11801',
    ('7858', '7837'): '305 WALTER AVE MINEOLA NY 11501',
    ('7702', '7704', '7710', '7715'): '3107 BAYLOR ST LUBBOCK TX 79415',
    ('7709', '7712'): '531 ROANE ST CHARLESTON WV 25302',
    ('7701', '7703', '7705', '7706', '7707', '7708', '7711', '7713', '7714'): '123 TEST TEST TEST TEST 12345'
}

markets_clusters: dict = {
    'optimum': {
        'markets': ['K', 'M', 'N', 'G'],
        'clusters': [6, 10, 86]
    },
    'suddenlink': {
        'markets': ['A', 'B', 'C', 'E', 'F', 'G', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'V'],
        'clusters': [10, 21, 58, 59, 66, 67, 90, 91, 92, 93, 95]
    }
}

def get_input_excel(event: gp.widgets.GooeyPieEvent) -> None:
    """
    get_input_excel Reads the input excel file.

    Reads the excel file uploaded and comverts the data into a dictionary that can be used for further processing.

    :param event: Reference of the widget/event that called this function.
    :type event: gp.widgets.GooeyPieEvent
    """
    global file_path, final_dict, data
    file_path = input_file_window.open()
    try:
        input_file_lbl.text = file_path.split('/')[-1]
    except AttributeError:
        print("No file was selected, ignoring error")
    else:
        input_file_lbl.color = 'green'
        data = pd.read_excel(file_path, skiprows=1, header=None, index_col=None, usecols='A:D', names=['ID', 'Gathering Name', 'Gathering Description', 'Gathering Price'])


def set_market_cluster(event: gp.widgets.GooeyPieEvent) -> None:
    """
    set_market_cluster Sets the market and cluster dropdopwns.

    Sets the market and cluster dropdopwn based on if OPT or SDL is selected as both the proposals have a different set of markets and clusters.

    :param event: Reference of the widget that called this function.
    :type event: gp.widgets.GooeyPieEvent
    """
    market_dd.items = markets_clusters[event.widget.selected.lower()]['markets']
    cluster_dd.items = markets_clusters[event.widget.selected.lower()]['clusters']
    
    eid_inp.disabled, ftax_inp.disabled, eid_lbl.disabled, ftax_lbl.disabled = (True, True, True, True) if event.widget.selected == 'Optimum' else (False, False, False, False)
    

def toggle_promo(event: gp.widgets.GooeyPieEvent) -> None:
    """
    toggle_promo Disables/enables the promo radio buttons.

    Disables or enables the promo radio buttons based on whether UOW is selected or not as UOW does not have the concept of full rate offers.

    :param event: Reference of the widget that called this function.
    :type event: gp.widgets.GooeyPieEvent
    """
    promo_rg.disabled = True if event.widget.selected == 'UOW' else False


def sanitize_corp(event: gp.widgets.GooeyPieEvent) -> None:
    """
    sanitize_corp Only allows the corp field to take in numerical values and a max length of 4 digits.

    Corp values are only 4 digits long if you exclude the leading 0 which is not required.
    This function makes sure the corp values are digits and that the max length doesn't exceed 4 characters.

    :param event: Reference of the widget that called this function
    :type event: gp.widgets.GooeyPieEvent
    """
    if event.widget.text and not event.widget.text[-1].isnumeric():
        event.widget.text = event.widget.text[:-1]
    if len(event.widget.text) > 4:
        event.widget.text = event.widget.text[:4]

    
def sanitize_ftax(event: gp.widgets.GooeyPieEvent) -> None:
    """
    sanitize_ftax Only allows the ftax field to take in numerical values and a max length of 2 digits.

    Ftax values are 2 digit identifiers and as such this function makes sure the user input is numerical and has a max length of 2 digits.

    :param event: Reference of the widget that called this function
    :type event: gp.widgets.GooeyPieEvent
    """
    if event.widget.text and not event.widget.text[-1].isnumeric():
        event.widget.text = event.widget.text[:-1]
    if len(event.widget.text) > 2:
        event.widget.text = event.widget.text[:2]
    

def sanitize_eid(event: gp.widgets.GooeyPieEvent) -> None:
    """
    sanitize_eid Only allows 5 characters and makes the entry uppercase.

    EIDs are alphanumerical identifiers of length 5, and so this function makes sure the length of the input is restricted to 5 and is converted to uppercase.

    :param event: Reference of the widget that called this function
    :type event: gp.widgets.GooeyPieEvent
    """
    event.widget.text = event.widget.text.upper()
    if len(event.widget.text) > 5:
        event.widget.text = event.widget.text[:5]


def validate_submit_values() -> None:
    """
    validate_submit_values Main driver function which does all the processing.

    Validates whether all inputs are provided and processes the data.
    1. Validates all input are present.
    2. Checks to see which request to make based on channel and env.
    3. Processes the response and appends the data to the dictionary created by using the input data.
    4. Compares the data and creates a new field with the result (Pass, Fail and NA).
    5. Saves the data to an excel.
    """
    if not file_path:
        app.alert("Error", "Please select a file!", "error")
        return
    if not all((proposal_rg.selected, channel_rg.selected, env_rg.selected, promo_rg.selected, corp_inp.text, market_dd.selected, cluster_dd.selected, ftax_inp.text or ftax_inp.disabled, eid_inp.text or eid_inp.disabled)):
        app.alert('Error', 'Please enter all values', 'error')
        return
    
    progress_bar.start(5)
    submit_btn.disabled = True
    app.update()
    
    final_dict.clear()
    for _, row in data.iterrows():
        final_dict[str(row['ID'])] = {
            'Gathering Name':row['Gathering Name'],
            'Gathering Description': row['Gathering Description'],
            'Gathering Price': f"{row['Gathering Price']:.2f}" if row['Gathering Price'] else ''
            }
    
    proposal = 'opt' if proposal_rg.selected == 'Optimum' else 'sdl'
    channel = channel_rg.selected.split('/')[-1].lower()
    env = env_rg.selected.lower()
    promo = 'true' if promo_rg.selected == 'Promotional' else 'false'
    corp = corp_inp.text
    market = market_dd.selected
    cluster = cluster_dd.selected
    ftax = ftax_inp.text
    eid = eid_inp.text
    
    addr_loc, addr_street, addr_city, addr_state, addr_zip = get_address(corp)
    url = urls[channel][env]
    
    match (channel, proposal):
        case ('uow', 'opt'):
            parameters = (cluster, corp, market, addr_loc, addr_street, addr_city, addr_state, addr_zip)
        case ('uow', 'sdl'):
            parameters = (eid, cluster, corp, ftax, market, addr_loc, addr_street, addr_city, addr_state, addr_zip)
        case ('dsa', 'opt'):
            parameters = (promo, cluster, market, corp)
        case ('dsa', 'sdl'):
            parameters = (promo, eid, ftax, cluster, market, corp)
    
    payload = payloads[channel][proposal].format(*parameters)
    request = json.loads(payload)
    
    if channel == 'uow':
        try:
            res = requests.post(url, json=request, auth=('unittest', 'test01'))
        except requests.exceptions.ConnectionError:
            handle_app_state_change_on_exceptions()
            app.alert("VPN Error", "Please connect to VPN or check your network connection!")
            return
        else:
            res = res.json()
            offers = res["productOfferings"]["productOfferingResults"]
    else:
        try:
            res = requests.post(url, json=request, verify=False)
        except requests.exceptions.ConnectionError:
            handle_app_state_change_on_exceptions()
            app.alert("VPN Error", "Please connect to VPN or check your network connection!")
            return
        else:
            res = res.json()
            offers = res["searchProductOfferingReturn"]["productOfferingResults"]
    
    if not offers:
        handle_app_state_change_on_exceptions()
        app.alert('Error', 'No offers returned, please use a different combination and try again', 'error')
        return
    
    with open('temp.json', 'w') as f:
        f.write(json.dumps(offers, indent=4))

    if mobile_offers_cb.checked:
        title = 'mobileTitle'
        description = 'mobileDescription'
        price = 'mobileDefaultPrice'
    else:
        title = 'title'
        description = 'description'
        price = 'defaultPrice'
    
    for offer in offers:
        if offer['matchingProductOffering']['ID'] in final_dict.keys():
            final_dict[offer['matchingProductOffering']['ID']].update({
                "EPC Name": offer['matchingProductOffering'][title],
                "EPC Description": offer['matchingProductOffering'][description],
                "EPC Price": f"{float(offer['matchingProductOffering'][price].split(':')[1]):.2f}" if offer['matchingProductOffering'][price] else ''
            })
    
    
    pass_list: list = []
    fail_list: list = []
    na_list: list = []
    
    for offer, attributes in final_dict.items():
        try:
            name_check = attributes['Gathering Name'] == attributes['EPC Name']
            description_check = (attributes['Gathering Description'] == attributes['EPC Description']) or not description_cb.checked
            price_check = (attributes['Gathering Price'] == attributes['EPC Price']) or not price_cb.checked
        except KeyError:
            temp = [offer, *attributes.values(), 'Not found', 'Not found', 'Not found', 'NA']
            na_list.append(temp)
        else:
            if all((name_check, description_check, price_check)):
                temp = [offer, *attributes.values(), 'Pass']
                pass_list.append(temp)
            else:
                temp = [offer, *attributes.values(), 'Fail']
                fail_list.append(temp)
    
    if proposal == 'opt':
        name_str = f'[Corp - {corp}][Market - {market}][Cluster - {cluster}]'
    else:
        name_str = f'[Corp - {corp}][Market - {market}][Cluster - {cluster}][Ftax - {ftax}][EID - {eid}]'
        
    save_excel(f'Pass {name_str}.xlsx', pass_list)
    save_excel(f'Fail {name_str}.xlsx', fail_list)
    save_excel(f'NA {name_str}.xlsx', na_list)
    
    handle_app_state_change_on_exceptions()
    result_tbl.clear()

    result_tbl.add_row([len(final_dict), len(pass_list), len(fail_list), len(na_list)])
    result_window.show_on_top()
    return


def handle_app_state_change_on_exceptions() -> None:
    """
    handle_app_state_change_on_exceptions Helper function to handle UI when exceptions occur.

    Handles the UI by stopping the Progressbar and enabling the button if an exception occurs anywhere in the main code.
    """
    progress_bar.stop()
    progress_bar.value = 0
    submit_btn.disabled = False
    app.update()


def save_excel(name: str, data: list):
    """
    save_excel Helper function to save data to an excel file.
    
    Takes in a name and list of lists and saves it to an excel file.
    
    :param name: name of the excel file to be saved.
    :type name: str
    
    :param data: The table data to be saved to excel file.
    :type data: List of lists
    """
    global headers
    df: pd.DataFrame = pd.DataFrame(data, columns=headers)
    try:
        with pd.ExcelWriter(name, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
    except PermissionError:
        handle_app_state_change_on_exceptions()
        app.alert('File open', 'The file where the output is being written is open.\nPlease close the file and try again!', 'warning')
    

def get_address(corp: str) -> tuple:
    """
    get_address Splits the address.

    Splits the address as required by the request payload.

    :param corp: Corp used while running the tool.
    :type corp: str
    :return: Tuple of strings with the different address attributes.
    :rtype: Tuple of strings
    """
    for k, v in corps.items():
        if corp in k:
            address = v.split()
            return address[0], ' '.join(address[1:3]), address[3], address[4], address[5]

# sys.excepthook = handle_exceptions
tk.Tk.report_callback_exception = handle_exceptions
threading.excepthook = handle_thread_exception

if __name__ == '__main__':
    app = gp.GooeyPieApp(f'Name Description Checker {version}')
    app.on_open(lambda: threading.Thread(target=check_version).start())

    input_file_window = gp.OpenFileWindow(app, 'Select input file')
    input_file_window.set_initial_folder('app')
    input_file_window.add_file_type("Excel files", '*.xlsx')

    input_file_btn = gp.Button(app, 'Select file', get_input_excel)
    input_file_lbl = gp.StyleLabel(app, 'No file selected')
    input_file_lbl.color = 'red'

    proposal_rg = gp.LabelRadiogroup(app, 'Proposal', ['Optimum', 'Suddenlink'], 'horizontal')
    proposal_rg.add_event_listener('change', set_market_cluster)

    channel_rg = gp.LabelRadiogroup(app, 'Channel', ['ISA/DSA', 'UOW'], 'horizontal')
    channel_rg.add_event_listener('change', toggle_promo)

    env_rg = gp.LabelRadiogroup(app, 'Environment', ['UAT', 'UAT1', 'UAT2'], 'horizontal')

    promo_rg = gp.LabelRadiogroup(app, 'Promotion', ['Promotional', 'Full Rate'], 'horizontal')
    promo_rg.selected_index = 0

    corp_container = gp.Container(app)

    corp_lbl = gp.Label(corp_container, 'Corp')
    corp_inp = gp.Input(corp_container)
    corp_inp.add_event_listener('change', sanitize_corp)

    container = gp.Container(app)

    market_lbl = gp.Label(container, 'Market')
    market_dd = gp.Dropdown(container, [])

    cluster_lbl = gp.Label(container, 'Cluster')
    cluster_dd = gp.Dropdown(container, [])

    ftax_lbl = gp.Label(container, 'Ftax')
    ftax_inp = gp.Input(container)
    ftax_inp.add_event_listener('change', sanitize_ftax)

    eid_lbl = gp.Label(container, 'EID')
    eid_inp = gp.Input(container)
    eid_inp.add_event_listener('change', sanitize_eid)

    checkbox_container = gp.LabelContainer(app, 'Parameters to Check')

    offer_id_cb = gp.Checkbox(checkbox_container, 'Offer ID', True)
    offer_id_cb.disabled = True

    name_cb = gp.Checkbox(checkbox_container, 'Name', True)
    name_cb.disabled = True

    description_cb = gp.Checkbox(checkbox_container, 'Description')
    price_cb = gp.Checkbox(checkbox_container, 'Price')

    mobile_offers_container = gp.Container(checkbox_container)
    mobile_offers_cb = gp.Checkbox(mobile_offers_container, 'Checking Mobile Offers?')

    submit_btn = gp.Button(checkbox_container, 'Submit', lambda x: threading.Thread(target=validate_submit_values).start())
    output_folder_btn = gp.Button(checkbox_container, 'Output Folder', lambda x: os.startfile(os.getcwd()))
    submit_btn.width = output_folder_btn.width
    
    progress_bar = gp.Progressbar(app, 'indeterminate')
    
    result_window = gp.Window(app, 'Result')
    
    result_header = gp.StyleLabel(result_window, 'Run statistics')
    result_header.font_size = 15
    result_header.font_weight = 'bold'
    
    result_tbl = gp.Table(result_window, ['Total', 'Pass', 'Fail', 'NA'])
    result_tbl.set_column_widths(100, 100, 100, 100)
    result_tbl.height = 2
    result_tbl.set_column_alignments('center', 'center', 'center', 'center')
    
    result_close_btn = gp.Button(result_window, 'Close', lambda x: result_window.hide())
    
    result_window.set_grid(3, 1)

    result_window.add(result_header, 1, 1, align='center')
    result_window.add(result_tbl, 2, 1)
    result_window.add(result_close_btn, 3, 1, align='center')
    
    result_window.hide()
        
    app.set_grid(9, 4)

    app.add(input_file_btn, 1, 1, column_span=4, align='center')
    app.add(input_file_lbl, 2, 1, column_span=4, align='center')

    app.add(proposal_rg, 3, 1, column_span=2, fill=True)

    app.add(channel_rg, 3, 3, column_span=2, fill=True)

    app.add(env_rg, 4, 1, column_span=2, fill=True)

    app.add(promo_rg, 4, 3, column_span=2, fill=True)

    app.add(corp_container, 5, 1, column_span=4, fill=True)

    app.add(progress_bar, 9, 2, column_span=2, align='center')
        
    corp_container.set_grid(1, 5)

    corp_container.add(corp_lbl, 1, 2, align='right')
    corp_container.add(corp_inp, 1, 3)

    app.add(container, 6, 1, column_span=4, row_span=2, fill=True)

    container.set_grid(2, 4)

    container.add(market_lbl, 1, 1, fill=True)
    container.add(market_dd, 1, 2, fill=True)

    container.add(cluster_lbl, 1, 3, fill=True)
    container.add(cluster_dd, 1, 4, fill=True)

    container.add(ftax_lbl, 2, 1, fill=True)
    container.add(ftax_inp, 2, 2, fill=True)

    container.add(eid_lbl, 2, 3, fill=True)
    container.add(eid_inp, 2, 4, fill=True)

    app.add(checkbox_container, 8, 1, column_span=4, fill=True)

    checkbox_container.set_grid(4, 4)

    checkbox_container.add(offer_id_cb, 1, 1, fill=True)
    checkbox_container.add(name_cb, 1, 2, fill=True)
    checkbox_container.add(description_cb, 1, 3, fill=True)
    checkbox_container.add(price_cb, 1, 4, fill=True)
    checkbox_container.add(mobile_offers_container, 2, 1, fill=True, column_span=4)
    checkbox_container.add(submit_btn, 3, 2, align='center')
    checkbox_container.add(output_folder_btn, 3, 3, align='center')

    mobile_offers_container.set_grid(1, 1)
    mobile_offers_container.add(mobile_offers_cb, 1, 1, align='center')

    app.run()
