import imaplib
import email
from bs4 import BeautifulSoup
import pandas as pd
from tqdm import tqdm

def extract_form_data(body):
    soup = BeautifulSoup(body, 'html.parser')
    data = {}
    for strong_tag in soup.find_all('strong'):
        label = strong_tag.get_text().strip()
        value = strong_tag.find_next_sibling(text=True).strip()
        data[label] = value
    return data

def calculate_acertos(data, correct_answers):
    acertos = 0
    for i in range(1, 11):
        question = f'Pergunta{i}'
        if question in data and data[question] == correct_answers[i-1]:
            acertos += 1
    return acertos

correct_answers = [
    'O continente resiste aos desejos usando a razão',  # Pergunta 1
    'Agir contra o próprio julgamento racional',        # Pergunta 2
    'Os prazeres legítimos são diferentes dos excessivos', # Pergunta 3
    'Manter o controle racional contra impulsos',        # Pergunta 4
    'O vicioso escolhe o mal conscientemente',           # Pergunta 5
    'Luxúria',                                           # Pergunta 6
    'Influencia a formação dos hábitos',                 # Pergunta 7
    'São complementares',                                # Pergunta 8
    'A virtude é adquirida através do hábito',           # Pergunta 9
    'Felicidade'                                         # Pergunta 10
]

mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login('testeste054@gmail.com', 'aipz hqqs fihw oskr')
mail.select('inbox')

status, data = mail.search(None, '(FROM "info@staticforms.xyz")')

email_ids = data[0].split()

form_data_list = []

for email_id in tqdm(email_ids, desc="Processando e-mails", unit="email"):
    status, msg_data = mail.fetch(email_id, '(RFC822)')
    msg = email.message_from_bytes(msg_data[0][1])
    
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                html_body = part.get_payload(decode=True).decode('utf-8')
                form_data = extract_form_data(html_body)
                if form_data:
                    form_data['Acertos'] = calculate_acertos(form_data, correct_answers)
                    form_data_list.append(form_data)
    else:
        if msg.get_content_type() == 'text/html':
            html_body = msg.get_payload(decode=True).decode('utf-8')
            form_data = extract_form_data(html_body)
            if form_data:
                form_data['Acertos'] = calculate_acertos(form_data, correct_answers)
                form_data_list.append(form_data)

mail.logout()

df = pd.DataFrame(form_data_list)

df_filtered = df[['Nome', 'Curso', 'Acertos']]

df_filtered.to_csv('form_submissions.csv', index=False)

print(df_filtered)
print("Dados extraídos e salvos com sucesso.")
