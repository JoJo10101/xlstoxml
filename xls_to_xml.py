#conding:utf-8
import xlrd
import xml.dom.minidom
import tkFileDialog


filename = tkFileDialog.askopenfilename(initialdir ='')


try:
	wkb=xlrd.open_workbook(filename)
	sheet=wkb.sheets()[0]

	nrows=sheet.nrows  

	doc = xml.dom.minidom.Document() 

	root = doc.createElement('testcases') 

	doc.appendChild(root) 

	for i in range (1,nrows):

		nodeTestcase = doc.createElement('testcase')

		nodeTestcase.setAttribute('internalid',str(i))
		nodeTestcase.setAttribute('name',sheet.cell(i,1).value.encode('utf-8'))

		nodeOrder = doc.createElement('node_order')
		nodeOrder.appendChild(doc.createTextNode('pythontest'))

		nodeExternalid = doc.createElement('externalid')
		nodeExternalid.appendChild(doc.createTextNode('pythontest'))

		nodeVersion = doc.createElement('version')
		nodeVersion.appendChild(doc.createTextNode('pythontest'))

		nodeSummary = doc.createElement('summary')
		nodeSummary.appendChild(doc.createTextNode(sheet.cell(i,3).value.encode('utf-8')))

		nodePreconditions = doc.createElement('preconditions')
		nodePreconditions.appendChild(doc.createTextNode(sheet.cell(i,4).value.encode('utf-8')))

		nodeExecution_type = doc.createElement('execution_type')
		nodeExecution_type.appendChild(doc.createTextNode('pythontest'))

		nodeImportance = doc.createElement('importance')
		nodeImportance.appendChild(doc.createTextNode(sheet.cell(i,2).value.encode('utf-8')))

		nodeSteps = doc.createElement('steps')

		nodeStep = doc.createElement('step')
		
		nodeStepnumber = doc.createElement('step_number')
		nodeStepnumber.appendChild(doc.createTextNode('1'))

		nodeActions = doc.createElement('actions')
		nodeActions.appendChild(doc.createTextNode(sheet.cell(i,5).value.encode('utf-8')))

		nodeExpectedresults = doc.createElement('expectedresults')
		nodeExpectedresults.appendChild(doc.createTextNode(sheet.cell(i,6).value.encode('utf-8')))

		nodeExecution_type = doc.createElement('execution_type')
		nodeExecution_type.appendChild(doc.createTextNode('1'))

		nodeSteps.appendChild(nodeStep)
		nodeStep.appendChild(nodeStepnumber)
		nodeStep.appendChild(nodeActions)
		nodeStep.appendChild(nodeExpectedresults)
		nodeStep.appendChild(nodeExecution_type)

		nodeTestcase.appendChild(nodeOrder)
		nodeTestcase.appendChild(nodeExternalid)
		nodeTestcase.appendChild(nodeVersion)
		nodeTestcase.appendChild(nodeSummary)
		nodeTestcase.appendChild(nodePreconditions)
		nodeTestcase.appendChild(nodeExecution_type)
		nodeTestcase.appendChild(nodeImportance)
		nodeTestcase.appendChild(nodeSteps)
		root.appendChild(nodeTestcase)


	try:

		fp = open(filename.replace('.xls','.xml'), 'w')
		doc.writexml(fp, indent='\t', addindent='\t', newl='\n', encoding="utf-8")

	except:

		print 'excel return wrong'

except:
	print 'excel is wrong'







