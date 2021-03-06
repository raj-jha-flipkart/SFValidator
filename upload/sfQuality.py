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

import cloudstorage as gcs
from google.appengine.api import app_identity
from google.appengine.api import mail
sender_email_id="raj.jha@flipkart.com"
########################################################################################################################################
UserID = "user id"
JobGrade = "job grade"
ReportingManagerID = "reporting manager id"
ProductionTemplateFileName = settings.PRODUCTION_TEMPLATE
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
	tsv_file ="/sfuploadvalidator.appspot.com/"+ ReportName + ".tsv"
	xlsx_file ="/sfuploadvalidator.appspot.com/"+ ReportName + ".xlsx"


	tsvHandle = gcs.open(tsv_file,"w")
	tsvHandle.write("User ID\tError\n")
	for error in Errors:
		tsvHandle.write(error+"\n")
	tsvHandle.close()
	return tsv_file
	
	workbook = Workbook(xlsx_file)
	worksheet = workbook.add_worksheet()
	
	tsv_reader = csv.reader(gcs.open(tsv_file, 'r'), delimiter='\t')

	for row, data in enumerate(tsv_reader):
		worksheet.write_row(row, 0, data)
	workbook.close()
	gcs.delete(tsv_file)

def XlsxToTsv(FilePath):
	reload(sys)
	sys.setdefaultencoding('utf-8')
	wb = xlrd.open_workbook(file_contents = gcs.open(FilePath).read())
	sh = wb.sheet_by_index(0)

	FilePathTsv= FilePath.split(".xlsx")[0]
	csvFile = gcs.open(FilePathTsv+'.tsv', 'w')
	wr = csv.writer(csvFile,delimiter='\t')

	for rownum in xrange(sh.nrows):
        	wr.writerow(sh.row_values(rownum))
	csvFile.close()
	gcs.delete(FilePath)
	return FilePathTsv+".tsv"


def Check(EmployeeDataFilePath,emailID,fallOutReport):
	FileName = EmployeeDataFilePath.split('/')[-1].split(".xlsx")[0]
	EmployeeDataFilePathTsv = XlsxToTsv(EmployeeDataFilePath)
	FieldId,Employees,TotalEmployee = EmployeeData(EmployeeDataFilePathTsv)
	Errors = ErrorList(Employees,TotalEmployee,ProductionTemplateFileName)
	if len(Errors) == 0 :
		if fallOutReport == 1:
			FileName = FallOutReportXlsx(FieldId,Employees,TotalEmployee,ProductionTemplateFileName,FileName)
			Message = open("SuccessEmailBody.txt").read()
			Subject = "Success Factor Upload FallOut Report"
			mail.send_mail(sender=sender_email_id.format(
                	app_identity.get_application_id()),
                	to=emailID,
                	subject=Subject,
                	body=Message,attachments=[(FileName, gcs.open(FileName).read())])
			gcs.delete(FileName)
	else:
		FileName = XlsxErrorReport(Errors,FileName)
		Message = open("ErrorEmailBody.txt").read()
		Subject = "Success Factor Upload File Error"
		mail.send_mail(sender=sender_email_id.format(
                app_identity.get_application_id()),
                to=emailID,
                subject=Subject,
                body=Message,attachments=[(FileName, gcs.open(FileName).read())])
		gcs.delete(FileName)
	if len(Errors) == 0:
		return 1
	else:
		return 0
	

	
