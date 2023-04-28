import gooeypie as gp


urls = {
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

corps = {
    ('7801', '7816'): '61 SLEEPY LN HICKSVILLE NY 11801',
    ('7858', '7837'): '305 WALTER AVE MINEOLA NY 11501',
    ('7702', '7704', '7710', '7715'): '3107 BAYLOR ST LUBBOCK TX 79415',
    ('7709', '7712'): '531 ROANE ST CHARLESTON WV 25302'
}

markets_clusters = {
    'optimum': {
        'markets': ['K', 'M', 'N', 'G'],
        'clusters': [6, 10, 86]
    },
    'suddenlink': {
        'markets': ['A', 'B', 'C', 'E', 'F', 'G', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'V'],
        'clusters': [10, 21, 58, 59, 66, 67, 90, 91, 92, 93, 95]
    }
}

file = None
offers = {}

def get_input_excel(event):
    global file, offers
    offers.clear()
    file = input_file_window.open()
    try:
        input_file_lbl.text = file.split('/')[-1]
    except AttributeError:
        print("No file was selected, ignoring error")
    else:
        input_file_lbl.color = 'green'
        
        


def set_market_cluster(event):
    market_dd.items = markets_clusters[event.widget.selected.lower()]['markets']
    cluster_dd.items = markets_clusters[event.widget.selected.lower()]['clusters']
    
    eid_inp.disabled = True if event.widget.selected == 'Optimum' else False
    ftax_inp.disabled = True if event.widget.selected == 'Optimum' else False
    

def toggle_promo(event):
    promo_rg.disabled = True if event.widget.selected == 'UOW' else False


# This is to only allow corp and ftax fields to accept numerical values
def sanitize_corp(event):
    if event.widget.text and not event.widget.text[-1].isnumeric():
        event.widget.text = event.widget.text[:-1]
    if len(event.widget.text) > 4:
        event.widget.text = event.widget.text[:4]

    
def sanitize_ftax(event):
    if event.widget.text and not event.widget.text[-1].isnumeric():
        event.widget.text = event.widget.text[:-1]
    if len(event.widget.text) > 2:
        event.widget.text = event.widget.text[:2]
    

def sanitize_eid(event):
    event.widget.text = event.widget.text.upper()
    if len(event.widget.text) > 5:
        event.widget.text = event.widget.text[:5]

def validate_submit_values(event):
    if not all((file,
          proposal_rg.selected,
          channel_rg.selected,
          env_rg.selected,
          promo_rg.selected,
          corp_inp.text,
          market_dd.selected,
          cluster_dd.selected,
          ftax_inp.text or ftax_inp.disabled,
          eid_inp.text or eid_inp.disabled)
          ):
        app.alert('Error', 'Please enter all values', 'error')
        return
    print(f"{file=}\n{proposal_rg.selected=}\n{channel_rg.selected=}\n{env_rg.selected=}\n{promo_rg.selected=}\n{corp_inp.text=}\n{market_dd.selected=}\n{cluster_dd.selected=}\n{ftax_inp.text or ftax_inp.disabled=}\n{eid_inp.text or eid_inp.disabled=}")


app = gp.GooeyPieApp('Offer Name Description Price Checker')

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

env_rg = gp.LabelRadiogroup(app, 'Environment (Middleware)', ['UAT', 'UAT1', 'UAT2'], 'horizontal')

promo_rg = gp.LabelRadiogroup(app, 'Promotion', ['Promo', 'Full Rate'], 'horizontal')

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

submit_btn = gp.Button(checkbox_container, 'Submit', validate_submit_values)

app.set_grid(8, 4)

app.add(input_file_btn, 1, 1, column_span=4, align='center', )
app.add(input_file_lbl, 2, 1, column_span=4, align='center', )

app.add(proposal_rg, 3, 1, column_span=2, fill=True)

app.add(channel_rg, 3, 3, column_span=2, fill=True)

app.add(env_rg, 4, 1, column_span=2, fill=True)

app.add(promo_rg, 4, 3, column_span=2, fill=True)

app.add(corp_container, 5, 1, column_span=4, fill=True)

corp_container.set_grid(1, 5)

corp_container.add(corp_lbl, 1, 2, align='right')
corp_container.add(corp_inp, 1, 3)

app.add(container, 6, 1, column_span=4, row_span=2)

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

checkbox_container.set_grid(3, 4)

checkbox_container.add(offer_id_cb, 1, 1, fill=True)
checkbox_container.add(name_cb, 1, 2, fill=True)
checkbox_container.add(description_cb, 1, 3, fill=True)
checkbox_container.add(price_cb, 1, 4, fill=True)
checkbox_container.add(mobile_offers_container, 2, 1, fill=True, column_span=4)
checkbox_container.add(submit_btn, 3, 2, column_span=2, align='center')

mobile_offers_container.set_grid(1, 1)
mobile_offers_container.add(mobile_offers_cb, 1, 1, align='center')

app.run()
