from calendar import calendar
from pprint import pprint
import pickle
import os
import datetime
from collections import namedtuple
from sys import api_version
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request




#É utilizada a API do Google, disponível com a assinatura do Google Cloud. Para utilizar a API do google estou utilizando a API desenvolvida por Jie Jenn https://learndataanalysis (guia do youtube de como utilizar a API Google Calendar)
def Create_Service(client_secret_file, api_name, api_version, *scopes, prefix=''):
	CLIENT_SECRET_FILE = client_secret_file
	API_SERVICE_NAME = api_name
	API_VERSION = api_version
	SCOPES = [scope for scope in scopes[0]]
	
	cred = None
	working_dir = os.getcwd()
	token_dir = 'token files'
	pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}{prefix}.pickle'

	### Check if token dir exists first, if not, create the folder
	if not os.path.exists(os.path.join(working_dir, token_dir)):
		os.mkdir(os.path.join(working_dir, token_dir))

	if os.path.exists(os.path.join(working_dir, token_dir, pickle_file)):
		with open(os.path.join(working_dir, token_dir, pickle_file), 'rb') as token:
			cred = pickle.load(token)

	if not cred or not cred.valid:
		if cred and cred.expired and cred.refresh_token:
			cred.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
			cred = flow.run_local_server()

		with open(os.path.join(working_dir, token_dir, pickle_file), 'wb') as token:
			pickle.dump(cred, token)

	try:
		service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
		print(API_SERVICE_NAME, API_VERSION, 'service created successfully')
		return service
	except Exception as e:
		print(e)
		print(f'Failed to create service instance for {API_SERVICE_NAME}')
		os.remove(os.path.join(working_dir, token_dir, pickle_file))
		return None

def convert_to_RFC_datetime(year=1900, month=1, day=1, hour=0, minute=0):
	dt = datetime.datetime(year, month, day, hour, minute, 0).isoformat() + 'Z'
	return dt

class GoogleSheetsHelper:
	# --> spreadsheets().batchUpdate()
	Paste_Type = namedtuple('_Paste_Type', 
					('normal', 'value', 'format', 'without_borders', 
					 'formula', 'date_validation', 'conditional_formatting')
					)('PASTE_NORMAL', 'PASTE_VALUES', 'PASTE_FORMAT', 'PASTE_NO_BORDERS', 
					  'PASTE_FORMULA', 'PASTE_DATA_VALIDATION', 'PASTE_CONDITIONAL_FORMATTING')

	Paste_Orientation = namedtuple('_Paste_Orientation', ('normal', 'transpose'))('NORMAL', 'TRANSPOSE')

	Merge_Type = namedtuple('_Merge_Type', ('merge_all', 'merge_columns', 'merge_rows')
					)('MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS')

	Delimiter_Type = namedtuple('_Delimiter_Type', ('comma', 'semicolon', 'period', 'space', 'custom', 'auto_detect')
						)('COMMA', 'SEMICOLON', 'PERIOD', 'SPACE', 'CUSTOM', 'AUTODETECT')

	# --> Types
	Dimension = namedtuple('_Dimension', ('rows', 'columns'))('ROWS', 'COLUMNS')

	Value_Input_Option = namedtuple('_Value_Input_Option', ('raw', 'user_entered'))('RAW', 'USER_ENTERED')

	Value_Render_Option = namedtuple('_Value_Render_Option',["formatted", "unformatted", "formula"]
							)("FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA")
                            
	@staticmethod
	def define_cell_range(
		sheet_id, 
		start_row_number=1, end_row_number=0, 
		start_column_number=None, end_column_number=0):
		"""GridRange object"""
		json_body = {
			'sheetId': sheet_id,
			'startRowIndex': start_row_number - 1,
			'endRowIndex': end_row_number,
			'startColumnIndex': start_column_number - 1,
			'endColumnIndex': end_column_number
		}
		return json_body

	@staticmethod
	def define_dimension_range(sheet_id, dimension, start_index, end_index):
		json_body = {
			'sheetId': sheet_id,
			'dimension': dimension,
			'startIndex': start_index,
			'endIndex': end_index
		}
		return json_body



class GoogleCalendarHelper:
	...

class GoogleDriverHelper:
	...

if __name__ == '__main__':
	g = GoogleSheetsHelper()
	print(g.Delimiter_Type)












#criação do token de autenticação. A API do Google cria um JSON  baixado diretamente pelo https://console.cloud.google.com/apis/credentials/ com os dados para autenticação, e o código abaixo cria o token de autenticação (pickling).
CLIENT_SECRET_FILE = 'client_secret_486321458503-bgc43a93mtvvq9rjll7hcb6pfj3v8nqh.apps.googleusercontent.com.json'
API_NAME = 'calendar'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/calendar']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

#cria um objeto que vai ser utilizado como parametro na criação do calendário https://developers.google.com/calendar/api/v3/reference/calendars/insert


sala_request_body = {
    #título do "calendário" - a sala.
    'summary':'Laboratório de Informática I',
	'description':'Sala com projetor utilizada para a prática de programação. São disponíveis 30 computadores de alto desempenho e mesa para utilização de notebooks pessoais'
}

def criaCalendario():
    response = service.calendars().insert(body=sala_request_body).execute()
    print(response)

#deleta calendário https://developers.google.com/calendar/api/v3/reference/calendars/delete
def deletaCalendarioId():
    service.calendars().delete(calendarId='6a8hs9l4hgh7iopm6se32vd66g@group.calendar.google.com').execute()



#https://developers.google.com/calendar/api/v3/reference/calendars/get
def listaSumarios():
    response = service.calendarList().list().execute()
    print('lista de calendários: \/\n')
    #pprint faz a quebra de linha
    pprint(response)
    print(response.keys())
    print('\n\n')
    #item é como se fosse o número do calendário
    #print('id do item 0:')
    #print(response.get('items')[0]['id'])
    item = 0
    while (item < len(response.get('items'))):
        print('\nTítulo: '+response.get('items')[item]['summary'] + '\nId: '+response.get('items')[item]['id'])
		#vê se o calendário tem descrição antes de listar para não dar erro
        if('description' in response.get('items')[item].keys()):
            print('Descrição: '+ response.get('items')[item]['description'])
        item = item+1

def listaSummaryCalendarioEspecifico(id):
    #response = service.calendarList().list().execute()
    #pprint(response.summary)
    calendar_list_entry = service.calendarList().get(calendarId= id).execute()
    print(calendar_list_entry['summary'])

#para criar o evento (que seria a reserva da sala no horário específico) é necessário informar o Id do calendário (que seria a Sala)




#https://developers.google.com/calendar/api/v3/reference/events/insert
#nested request body https://developers.google.com/calendar/api/v3/reference/events/insert#examples
reserva_request_body = {
	#disciplina
	'summary': 'Aula de Programação IV',
	#conteúdo da aula
	'description': 'API REST',
	#pessoas que vão ter acesso ao evento - professores: https://developers.google.com/calendar/api/concepts/sharing
	'atendees':[
		{
			'email':'mas@aluno.ueg.br',
			#esta pessoa vai ter permissão para editar
			'organizer':True
		}
	],
	#documento do evento (aula)
	#'attachments': [
		#{
		#	'fileUrl':'https://docs.google.com/presentation/d/1QE5FXK85JUBlaPqoiFzMm4asiOw-Zc9TC81JIm70cTU/edit?usp=sharing',
		#	'tittle':'Slides da aula - Requisitos de Software SWECOM'
	#	}
	#],
	'start': {
    	'dateTime': '2022-07-22T20:50:00-03:00'
	},
	'end': {
		'dateTime': '2022-07-22T22:30:00-03:00'
	}
}

#variáveis para criação da reserva
#mandar os slides da aula
supportsAttachment = True
#id do calendário (control C do terminal do método listarSumarios)
idDoCalendario = 'iq88i2ljpct8rq92m8ssnpp2is@group.calendar.google.com'
#criar a reserva

def criaReserva():
	response = service.events().insert(
		#precisa fornecer o Id do calendário
		calendarId=idDoCalendario,
		supportsAttachments=supportsAttachment,
		body=reserva_request_body
	).execute()

#def update():
#idDoCalendario = response['id'] - vai depender da response
#idDoEvento = response['id] - vai depender da response
#service.events().update(
#	calendarId=idDoCalendario,
#	reserva_body_request
#).execute()

idDoEvento = 'rop9hjpsn266hpqlgc2rikli28'
def excluirReserva():
	service.events().delete(calendarId=idDoCalendario, eventId=idDoEvento).execute()

#$https://developers.google.com/calendar/api/v3/reference/events/list?hl=en
#pra ver o horário da reserva foi necessário acessar uma chave de um dicicionário dentro de outro dicionário, obter uma fatia de string específica dessa chave e converter para o fuso horário brasileiro.
def listaReservas():
	response = service.events().list(calendarId=idDoCalendario).execute()
	#pprint(response)
	item = 0
	while (item < len(response.get('items'))):
		horaInicio = (response.get('items')[item]['start']['dateTime'][11:16])
#		horaInicio = str(int(horaInicio[0:2])-3)+horaInicio[2:5]
		
		horaTermino = response.get('items')[item]['end']['dateTime'][11:16]
#		horaTermino = str(int(horaTermino[0:2])-3)+horaTermino[2:5]
#		if(horaTermino=='-2:30'):
#			horaTermino='22:30'

		print('\nReserva: '+response.get('items')[item]['summary'] )
		print('Horário: das '+ str(horaInicio)+' às '+ str(horaTermino))
		#vê se o calendário tem descrição antes de listar para não dar erro
		if('description' in response.get('items')[item].keys()):
			print('Conteúdo: '+ response.get('items')[item]['description'])
		print('Id: '+response.get('items')[item]['id'])
		item = item+1

def main():
	criaReserva()
	

main()