from __future__ import print_function
import preProcess,xlrd,sys,os,csv
from xlsxwriter.workbook import Workbook
from employeeData import EmployeeData
from FallOutReport import FallOutReportXlsx
from django.contrib.staticfiles.templatetags.staticfiles import static
from django.core.mail import EmailMessage
from django.core import mail
from django.conf import settings

#from __future__ import print_function
import httplib2
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes



########################################################################################################################################
UserID = "user id"
JobGrade = "job grade"
ReportingManagerID = "reporting manager id"
ProductionTemplateFileName = settings.BASE_DIR+'/static/'+settings.PRODUCTION_TEMPLATE

SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'sfValidator'
#########################################################################################################################################

def ChangeEmployeeDataFilepath(FilePath):
	global EmployeeDatafilePath
	EmployeeDataFilePath = FilePath

def VerifyStructure(Employees,TotalEmployee,ProductionTemplateFileName):
        LevelTitle,Table= preProcess.HashTable(ProductionTemplateFileName)
        LevelTitle = preProcess.TrimArraySpace(LevelTitle)
        Result = []
	Index = []
        for idx in xrange(TotalEmployee):
                for i in xrange(1,len(LevelTitle)):
			
                        empStructure = Employees[LevelTitle[i]][idx]
                        if Employees[LevelTitle[i-1]][idx] not in Table[i-1].keys():
                                Result.append(Employees[UserID][idx])
                                break
                        actualStructure =  Table[i-1][Employees[LevelTitle[i-1]][idx]]
                        if empStructure not in actualStructure:
                                Result.append(Employees[UserID][idx] + "\t" + LevelTitle[i].title() + " " + "Error")		
                                break
        return Result
		
def UserIDtoJobGrade(Employees,TotalEmployee):
	Hash = {}
	for idx in xrange(TotalEmployee):
		Hash[Employees[UserID][idx]] = Employees[JobGrade][idx]
	return Hash
		
def ManagerDiscripencies(Employees,TotalEmployee):
	Result = []
	Runtime = []
	Hash = UserIDtoJobGrade(Employees,TotalEmployee)
	for idx in xrange(TotalEmployee):
		try:
			JobGradeEmployee = Hash[Employees[UserID][idx]] 
			JobGradeManager = Hash[Employees[ReportingManagerID][idx]]
			if int(JobGradeEmployee) > int(JobGradeManager) :	
				Result.append(Employees[UserID][idx] + "\t" + "Manager Code Error")
		except KeyError,  e:
			Runtime.append(Employees[UserID][idx] + "\t" + "Manager Code Error" )
		except ValueError, e:
			Runtime.append(Employees[UserID][idx] + "\t" + "Manager Code Error")
	return Result + Runtime

def ErrorList(Employees,TotalEmployee,ProductionTemplateFileName):
	Result = VerifyStructure(Employees,TotalEmployee,ProductionTemplateFileName)
	Result += ManagerDiscripencies(Employees,TotalEmployee)
	return list(set(Result))


def XlsxErrorReport(Errors,ReportName):
	tsv_file = ReportName + ".tsv"
	xlsx_file = ReportName + ".xlsx"


	tsvHandle = open(tsv_file,"w")
	tsvHandle.write("User ID\tError\n")
	for error in Errors:
		tsvHandle.write(error+"\n")
	tsvHandle.close()

	
	workbook = Workbook(xlsx_file)
	worksheet = workbook.add_worksheet()
	
	tsv_reader = csv.reader(open(tsv_file, 'rb'), delimiter='\t')

	for row, data in enumerate(tsv_reader):
		worksheet.write_row(row, 0, data)
	workbook.close()
	os.remove(tsv_file)

def XlsxToTsv(FilePath):
	reload(sys)
	sys.setdefaultencoding('utf-8')
	wb = xlrd.open_workbook(FilePath)
	sh = wb.sheet_by_index(0)

	FilePathTsv= FilePath.split(".")[0]
	csvFile = open(FilePathTsv+'.tsv', 'wu')
	wr = csv.writer(csvFile,delimiter='\t')

	for rownum in xrange(sh.nrows):
        	wr.writerow(sh.row_values(rownum))
	csvFile.close()
	os.remove(FilePath)
	return FilePathTsv+".tsv"


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials	
def CreateMessage(Subject,Message,FileName,To_email):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEText(Message)
  message['to'] = To_email
  message['from'] = "me"
  message['subject'] = Subject
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def CreateMessageWithAttachment(Subject,Message,FileName,emailID):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file_dir: The directory containing the file to be attached.
    filename: The name of the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = emailID
  message['from'] = "me"
  message['subject'] = Subject

  msg = MIMEText(Message)
  message.attach(msg)

  path = FileName
  content_type, encoding = mimetypes.guess_type(path)

  if content_type is None or encoding is not None:
    content_type = 'application/octet-stream'
  main_type, sub_type = content_type.split('/', 1)
  if main_type == 'text':
    fp = open(path, 'rb')
    msg = MIMEText(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'image':
    fp = open(path, 'rb')
    msg = MIMEImage(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'audio':
    fp = open(path, 'rb')
    msg = MIMEAudio(fp.read(), _subtype=sub_type)
    fp.close()
  else:
    fp = open(path, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()

  msg.add_header('Content-Disposition', 'attachment', filename=FileName)
  message.attach(msg)

  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def SendMessage(Subject,Message,FileName,To_email):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.

  """
  credentials = get_credentials()
  http = credentials.authorize(httplib2.Http())
  service = discovery.build('gmail', 'v1', http=http)
  message = CreateMessageWithAttachment(Subject,Message,FileName,To_email)
  
  message = (service.users().messages().send(userId="me", body=message).execute())
  print ('Message Id: %s' % message['id'])
  return message
	
def Check(EmployeeDataFilePath,emailID,fallOutReport):
	FileName = EmployeeDataFilePath.split('/')[-1].split(".xlsx")[0]
	EmployeeDataFilePathTsv = XlsxToTsv(EmployeeDataFilePath)
	FieldId,Employees,TotalEmployee = EmployeeData(EmployeeDataFilePathTsv)
	Errors = ErrorList(Employees,TotalEmployee,ProductionTemplateFileName)
	if len(Errors) == 0 :
		if fallOutReport == 1:
			FallOutReportXlsx(FieldId,Employees,TotalEmployee,ProductionTemplateFileName,FileName)
			Message = open("SuccessEmailBody.txt").read()
			Subject = "Success Factor Upload FallOut Report"
			SendMessage(Subject,Message,FileName+".xlsx",emailID)
			os.remove(FileName+".xlsx")
	else:
		XlsxErrorReport(Errors,FileName)
		Message = open("ErrorEmailBody.txt").read()
		Subject = "Success Factor Upload File Error"
		SendMessage(Subject,Message,FileName+".xlsx",emailID)
		os.remove(FileName+".xlsx")
	if len(Errors) == 0:
		return 1
	else:
		return 0
	

	
