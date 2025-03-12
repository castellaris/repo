Dim objExcel, objWorkbook, objWorksheet, objRegression
Dim lastRow, i

' Открываем Excel и загружаем файл
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True ' Делаем Excel видимым
Set objWorkbook = objExcel.Workbooks.Open("C:\\Users\\User\\Desktop\\data.xlsx") ' Укажи путь к файлу
Set objWorksheet = objWorkbook.Sheets(1)

' Находим последнюю строку
lastRow = objWorksheet.Cells(objWorksheet.Rows.Count, 1).End(-4162).Row

' Выбираем диапазоны для регрессии
Dim yRange, xRange
Set yRange = objWorksheet.Range("B2:B" & lastRow) ' Зависимая переменная Y
Set xRange = objWorksheet.Range("C2:F" & lastRow) ' Независимые переменные X1-X4

' Запускаем инструмент регрессии
Set objRegression = objExcel.Application.AddIns("Analysis ToolPak").Installed
If Not objRegression Then
    objExcel.Application.AddIns("Analysis ToolPak").Installed = True
End If

objExcel.Run "ATPVBAEN.XLAM!Regress", yRange, xRange, True, True, , , , , True

' Закрываем Excel
objWorkbook.Save
objExcel.Quit

' Освобождаем память
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
WScript.Echo "Регрессионный анализ завершен!"
