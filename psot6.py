#Trabalho t6 PSO - João Gonçalves Machado 
import imaplib
import email
import os
import hashlib
import getpass
import os
from collections import defaultdict, Counter
import platform
from email.parser import HeaderParser


fileNameCounter = Counter()
fileNameHashes = defaultdict(set)
ProcessedMsgIDs = set()


def downloadAnexo(email_message, namefile):
    for part in email_message.walk():
        
        if part.get_content_maintype() == 'multipart': #multipart são containers
            #print(part.as_string())
            continue
        
        if part.get('Content-Disposition') is None:
            #print(part.as_string())
            continue
        
        file_name = part.get_filename() #pega o nome do arquivo anexo

        if file_name is not None:
            file_name = ''.join(file_name.splitlines())
        if file_name:
            payload = part.get_payload(decode=True)
            if payload:
                x_hash = hashlib.md5(payload).hexdigest()

               	#pula arquivos duplicados e arruma a nome da pasta e salva o anexo
                if x_hash in fileNameHashes[file_name]:
                    print('\tSkipping duplicate file: {file}'.format(file=file_name))
                    continue
                fileNameCounter[file_name] += 1
                file_str, file_extension = os.path.splitext(file_name)
                if fileNameCounter[file_name] > 1:
                    new_file_name = '{file}({suffix}){ext}'.format(suffix=fileNameCounter[file_name],
                                                                   file=file_str, ext=file_extension)
                    print('\tRenaming and storing: {file} to {new_file}'.format(file=file_name,
                                                                                new_file=new_file_name))
                else:
                    new_file_name = file_name
                    print('\tStoring: {file}'.format(file=file_name))
                fileNameHashes[file_name].add(x_hash)
                file_path = os.path.join(namefile, new_file_name)
                if os.path.exists(file_path):
                    print('\tExists in destination: {file}'.format(file=new_file_name))
                    continue
                try:
                    with open(file_path, 'wb') as fp:
                        fp.write(payload)
                except EnvironmentError:
                    print('Could not store: {file} it has a shitty file name or path under {op_sys}.'.format(
                        file=file_path,
                        op_sys=platform.system()))
            else:
                print('Attachment {file} was returned as type: {ftype} skipping...'.format(file=file_name,
                                                                                           ftype=type(payload)))
                continue




#conectando
mail = imaplib.IMAP4_SSL('imap.gmail.com')
password = getpass.getpass()
mail.login('jgmachado@inf.ufsm.br', password)
mail.list() #lista pastas gmail
mail.select("TrabalhoPSO") #conecta com uma pasta do email
sub = 'T6PSO' #este é o assunto do email

#data = ids dos emails
result, data = mail.search(None, "ALL")


ids = data[0]

#separando os ids em uma lista
id_list = ids.split()

#percorre a lista de emails
for msg in id_list[::-1]:
	print("procurando")
	result, data = mail.fetch(msg, '(RFC822)') #fetch do corpo do email com o id
	raw_email = data[0][1]
	email_message = email.message_from_bytes(raw_email)
	#se o assunto for igual ao esperado
	if(sub in email_message['Subject']):
		print ('Assunto: ' + email_message['Subject'])
		print("\n")
		print('De: ' + email_message['From'])
		print("\n\n")
		#se não existe a pasta do remetente cria
		if email_message['From'] not in os.listdir(os.getcwd()):
			os.mkdir(email_message['From'])
		#função para baixar o anexo
		downloadAnexo(email_message, email_message['From'])

