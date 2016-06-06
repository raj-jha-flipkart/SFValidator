#from preProcess import HashTable
import preProcess,xlrd,sys,os,csv
from xlsxwriter.workbook import Workbook
from employeeData import EmployeeData
from FallOutReport import FallOutReportXlsx
from django.contrib.staticfiles.templatetags.staticfiles import static
from django.core.mail import EmailMessage
from django.core import mail
from django.conf import settings


########################################################################################################################################
UserID = "user id"
JobGrade = "job grade"
ReportingManagerID = "reporting manager id"
ProductionTemplateFileName = settings.BASE_DIR+'/static/'+settings.PRODUCTION_TEMPLATE
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


def send_mail(Subject,Message,FileName,To_email):
	try:
		To_email = [To_email]
		email = EmailMessage(Subject,Message, To_email)
		email.attach_file(FileName)
		email.send(fail_silently=False)
	except:
		print "Mail was Not Send"

	
def Check(EmployeeDataFilePath,emailID,fallOutReport):
	FileName = EmployeeDataFilePath.split('/')[-1].split(".xlsx")[0]
	EmployeeDataFilePathTsv = XlsxToTsv(EmployeeDataFilePath)
	FieldId,Employees,TotalEmployee = EmployeeData(EmployeeDataFilePathTsv)
	Errors = ErrorList(Employees,TotalEmployee,ProductionTemplateFileName)
	if len(Errors) == 0 :
		if fallOutReport == 1:
			FallOutReportXlsx(FieldId,Employees,TotalEmployee,ProductionTemplateFileName,FileName)
			Message = ""
			Subject = "Success Factor Upload FallOut Report"
			send_mail(Subject,Message,FileName+".xlsx",emailID)
			os.remove(FileName+".xlsx")
	else:
		XlsxErrorReport(Errors,FileName)
		Message = "Error Message"
		Subject = "Success Factor Upload File Error"
		send_mail(Subject,Message,FileName+".xlsx",emailID)
		os.remove(FileName+".xlsx")
	if len(Errors) == 0:
		return 1
	else:
		return 0
	

	
