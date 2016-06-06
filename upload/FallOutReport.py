from employeeData import EmployeeData
import csv
import os
from xlsxwriter.workbook import Workbook

fileName = "ProductionTemplate.tsv"
class Node:
	def __init__(self):
		self.code ="Invalid"
		self.child = {}
def TreeStructure(fileName):
	File = open(fileName,"ru")
	LevelTitle = File.readline().split("\t")
	LevelTitle =[ LevelTitle[i].lower() for i in xrange(0,len(LevelTitle),2)]
	root = Node()
	Records = File.read().splitlines()
	for record in Records:
		record = record.split("\t")
		record = [record[i].lower() for i in xrange(len(record))]
		
		n = len(record)
		curNode = root
		for i in xrange(0,n,2):
			if record[i] in curNode.child.keys():
				curNode = curNode.child[record[i]]
			else:
				curNode.child[record[i]] = Node()
				curNode.child[record[i]].code= record[i+1]
				curNode = curNode.child[record[i]]
	return LevelTitle,root

def FallOutReport(Employees,TotalEmployee,fileName,ReportName):
	LevelTitle,root = TreeStructure(fileName)
	for idx in xrange(TotalEmployee):
		Handle = root
		for i in xrange(len(LevelTitle)):
			if LevelTitle[i] == "cost code":
				continue
			Key = Employees[LevelTitle[i]][idx]
			Employees[LevelTitle[i]][idx]= Handle.child[Key].code
			Handle = Handle.child[Key]
	return Employees

def FallOutReportXlsx(FieldId,Employees,TotalEmployee,fileName,ReportName):
	Employees = FallOutReport(Employees,TotalEmployee,fileName)
	tsvHandle = open("FallOutReport.tsv","w")
	for i in xrange(len(FieldId)):
		if i == len(FieldId)-1:
			tsvHandle.write(FieldId[i].title().strip("\n")+"\n")
		else:
        		tsvHandle.write(FieldId[i].title() + "\t")
        for idx in xrange(TotalEmployee):
		for i in xrange(len(FieldId)):
			if i == len(FieldId)-1:
				tsvHandle.write(Employees[FieldId[i]][idx].title()+"\n")
			else:
				tsvHandle.write(Employees[FieldId[i]][idx].title()+"\t")
        tsvHandle.close()

        tsv_file = ReportName + ".tsv"
        xlsx_file = ReportName + ".xlsx"

        workbook = Workbook(xlsx_file)
        worksheet = workbook.add_worksheet()

        tsv_reader = csv.reader(open(tsv_file, 'rb'), delimiter='\t')

        for row, data in enumerate(tsv_reader):
                worksheet.write_row(row, 0, data)
        workbook.close()

        os.remove(tsv_file)
	
if __name__ == "__main__":	
	FieldId,Employees,TotalEmployee = EmployeeData("Input/testData.tsv")
	FallOutReportXlsx(FieldId,Employees,TotalEmployee,fileName)
	
	
	
		
	
	

			
	
	


