from openpyxl import Workbook, load_workbook
wb = load_workbook('/var/lib/jenkins/workspace/Automation/Testsuite.xlsx')
ws = wb['TestSuite']
test_suite_values=[ws['A2'].value, ws['A3'].value,ws['A4'].value,
ws['A5'].value,ws['A6'].value, ws['A7'].value,ws['A8'].value,ws['A9'].value,ws['A10'].value,
ws['A11'].value,ws['A12'].value,ws['A13'].value,ws['A14'].value,ws['A15'].value,ws['A16'].value,
ws['A17'].value,ws['A18'].value,ws['A19'].value,ws['A20'].value,ws['A21'].value,ws['A22'].value,
ws['A23'].value,ws['A24'].value,ws['A25'].value,ws['A26'].value,ws['A27'].value,ws['A28'].value,
ws['A29'].value,ws['A30'].value]
test_suite_file_path=[ws['B2'].value, ws['B3'].value,ws['B4'].value,
ws['B5'].value,ws['B6'].value, ws['B7'].value,ws['B8'].value,ws['B9'].value,ws['B10'].value,
ws['B11'].value,ws['B12'].value,ws['B13'].value,ws['B14'].value,ws['B15'].value,ws['B16'].value,
ws['B17'].value,ws['B18'].value,ws['B19'].value,ws['B20'].value,ws['B21'].value,ws['B22'].value,
ws['B23'].value,ws['B24'].value,ws['B25'].value,ws['B26'].value,ws['B27'].value,ws['B28'].value,
ws['B29'].value,ws['B30'].value]
input_file_name= input("please input file name")
if input_file_name in test_suite_file_path:
    t1=test_suite_file_path.index(input_file_name)+2
    ws['A'+str(t1)].value=1
    print(ws['A'+str(t1)].value)
    wb.save('/var/lib/jenkins/workspace/Automation/Testsuite.xlsx')
wb2=load_workbook(input_file_name)
ws2 = wb2['Testcases']
if input_file_name=='/var/lib/jenkins/workspace/Automation/person_dim_test_suite.xlsx':
   input_product_name=input("please enter product name")
   if input_product_name=='telecom-dev':
      ws2['A2'].value=ws2['A3'].value=ws2['A4'].value=1
      wb2.save(input_file_name)
   elif input_product_name=="mobi-dev":
        ws2['A5'].value=ws2['A6'].value=ws2['A7'].value=1
        wb2.save(input_file_name)
   elif input_product_name=="rivermine-dev":
        ws2['A8'].value=ws2['A9'].value=ws2['A10'].value=ws2['A11'].value=1
        wb2.save(input_file_name)
   else:
    print('input product name is invalid')
else: exit

reset_values=input("Input True or false to reset all values")
if reset_values=="True":
    ws['A2'].value= ws['A3'].value=ws['A4'].value=ws['A5'].value=ws['A6'].value= ws['A7'].value=ws['A8'].value=ws['A9'].value=ws['A10'].value=ws['A11'].value=ws['A12'].value=ws['A13'].value=ws['A14'].value=ws['A15'].value=ws['A16'].value=ws['A17'].value=ws['A18'].value=ws['A19'].value=ws['A20'].value=ws['A21'].value=ws['A22'].value=ws['A23'].value=ws['A24'].value=ws['A25'].value=ws['A26'].value=ws['A27'].value=ws['A28'].value=ws['A29'].value=ws['A30'].value=0
    wb.save('/var/lib/jenkins/workspace/Automation/Testsuite.xlsx')
    ws2['A2'].value=ws2['A3'].value=ws2['A4'].value=ws2['A5'].value=ws2['A6'].value=ws2['A7'].value=ws2['A8'].value=ws2['A9'].value=ws2['A10'].value=ws2['A11'].value=0
    wb2.save(input_file_name)



