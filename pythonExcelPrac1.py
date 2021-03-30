import win32com.client
import pythoncom

# 서브 스레드에서 COM 객체를 사용하려면 COM 라이브러리를 초기화 해야함
pythoncom.CoInitialize()

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open(r'C:\python\pythonExcel\test2.xlsm')

# 새파일 만들기
# wb = excel.Workbooks.Add()

ws = wb.ActiveSheet

# worksheet를 추가
# ws2 = wb.Worksheets.Add()
# ws2.Name = "Number3"

print(ws.Cells(1,1).Value)

ws.Cells(1, 1).Value = "hello world" 
ws.Range('A1:B2').Value = "a"
ws.Range('A3:B4, A6:B7').Value = "b"

ws.Range("C1").Interior.ColorIndex = 10
ws.Range("A2:C2").Interior.ColorIndex = 20
ws.Range("C5").Font.ColorIndex = 1

# cell의 폭 조정
ws.Columns(1).ColumnWidth = 10
ws.Range("B:B").ColumnWidth = 20

# cell의 높이 조정
ws.Rows(1).RowHeight = 40
ws.Range("2:2").RowHeight = 60

# Macro 실행
excel.Application.Run("macro1")

# excel.SaveAs(r'C:\python\pythonExcel\test1.xlsx')
wb.Save()
# wb.Close(SaveChanges=False)
excel.Quit()

# 사용 후 uninitialize
pythoncom.CoUninitialize()
