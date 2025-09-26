import os, re, datetime, pickle, base64
import pandas as pd
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Gmail readonly scope
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# ---------- Authentication ----------
def auth_gmail():
    creds = None
    if os.path.exists('token.pickle'):
        creds = pickle.load(open('token.pickle','rb'))
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle','wb') as f:
            pickle.dump(creds, f)
    return build('gmail', 'v1', credentials=creds)

# ---------- Gmail Helpers ----------
def list_messages(service, query, max_results=50):
    resp = service.users().messages().list(userId='me', q=query, maxResults=max_results).execute()
    return resp.get('messages', [])

def get_message_text(service, msg_id):
    msg = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
    text = ''
    def walk(p):
        nonlocal text
        if 'body' in p and 'data' in p['body'] and p['body']['data']:
            txt = base64.urlsafe_b64decode(p['body']['data'].encode('UTF-8')).decode('utf-8', errors='ignore')
            text += txt + '\n'
        if 'parts' in p:
            for sp in p['parts']:
                walk(sp)
    walk(msg.get('payload', {}))
    return text, msg

# ---------- Parsing ----------
AMOUNT_RE = re.compile(r'paid\s*₹([0-9,.]+)', re.IGNORECASE)
PAYEE_RE = re.compile(r'paid\s*₹[0-9,.]+\s*to\s+([A-Za-z0-9 .,&-]+?)\s+at', re.IGNORECASE)
TIME_RE = re.compile(r'at\s+([0-9:APM ]+)\s+IST,\s+([0-9]{1,2}\s+\w+\s+[0-9]{4})', re.IGNORECASE)

def parse_transaction(text, platform='FamPay'):
    amount, payee, time_str = None, None, None
    
    m = AMOUNT_RE.search(text)
    if m: amount = m.group(1).replace(',', '')
    
    m = PAYEE_RE.search(text)
    if m: payee = m.group(1).strip()
    
    m = TIME_RE.search(text)
    if m:
        time_raw = m.group(1) + " " + m.group(2)  # "07:56 AM 15 September 2025"
        try:
            dt = datetime.datetime.strptime(time_raw, "%I:%M %p %d %B %Y")
            time_str = dt.strftime("%Y-%m-%d %H:%M")
        except:
            time_str = time_raw  # fallback
    
    return {
        'platform': platform,
        'amount': amount,
        'payee': payee,
        'description': 'FamPay Payment',
        'time': time_str
    }

# ---------- Storage ----------
def append_to_excel(row, fname='transactions.xlsx'):
    cols = ['SL no.', 'name of platform', 'how much paid', 'whom to paid', 'description', 'time']
    if os.path.exists(fname):
        df = pd.read_excel(fname)
    else:
        df = pd.DataFrame(columns=cols)
    
    # remove old total
    df = df[df['description'] != 'Total Spent']
    
    new_row = {
        'SL no.': len(df) + 1,
        'name of platform': row.get('platform'),
        'how much paid': float(row.get('amount') or 0),
        'whom to paid': row.get('payee'),
        'description': row.get('description'),
        'time': row.get('time')
    }
    
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    
    # add total row
    total = df['how much paid'].sum()
    total_row = {
        'SL no.': '',
        'name of platform': '',
        'how much paid': total,
        'whom to paid': '',
        'description': 'Total Spent',
        'time': ''
    }
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    
    df.to_excel(fname, index=False)

# ---------- Main ----------
def main():
    svc = auth_gmail()
    query = 'from:no-reply@famapp.in "Your payment of"'  # only FamApp receipts
    messages = list_messages(svc, query, max_results=50)
    
    for m in messages:
        mid = m['id']
        txt, msg = get_message_text(svc, mid)
        parsed = parse_transaction(txt, 'FamPay')
        if parsed['amount']:
            append_to_excel(parsed)

if __name__=='__main__':
    main()
