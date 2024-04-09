# VBA-challenge
VBA Challenge HW2 
Background 
In this homework assignment, you will use VBA scripting to analyze generated stock market data.
Objective 
The goal of this assignment is to create a VBA script that loops through all the stocks for one year and outputs to new columns, add conditional formatting, and calculate greatest yearly increase, decrease % and volumes and apply these to all worksheets in the workbook/Excel.

Dataset/file: 
Module 2 challenge files, Multiple Years Stock Data excel file that has columns in each sheet 
<ticker>	<date>	<open>	<high>	<low>	<close>	<vol>
…and file has 3 sheets 2018, 2019, and 2020.

Tasks 
1.	Create a script that loops through all the stocks for one year and outputs the following information:
a.	The ticker symbol
b.	Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
c.	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
d.	The total stock volume of the stock.
2.	Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
3.	Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

Deliverables 
•	Readme 
•	How to run the code: 
To execute this VBA project for analyzing stock market data, follow these steps:
1.	Set Up Repository:
o	Create a new repository named "VBA-challenge" for this project.
2.	Download Files:
o	Download the necessary files from Module 2 Challenge Files (MultipleYearStockData.xlsx)
3.	Open Excel and save as xlsm
o	Enable Developer Tab and If the Developer tab is not already enabled, enable it in Excel. 
4.	Open the Visual Basic for Applications (VBA) editor by clicking Record Macro (on Mac)
5.	Write VBA code 
6.	Select the macro created (e.g., analyzeStockData) and click Run to execute the script
7.	Review results   



•	Code
1.	Create a script that loops through all the stocks for one year and outputs the following information:
a.	The ticker symbol
b.	Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
c.	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
d.	The total stock volume of the stock.

StockData1.vba (StockDataAnalysis.vba in alphabetical_testing 21.xlsm)
Sub StockData1()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("2018") ' Change the sheet name as needed

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double

    ' Set headers for the output columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Dim summaryRow As Long
    summaryRow = 2

    For i = 2 To lastRow
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' New ticker symbol
            ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            totalVolume = 0
        End If

        ' Add the volume to the total volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Last row for the ticker symbol
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            If openingPrice <> 0 Then
                percentageChange = yearlyChange / openingPrice
            Else
                percentageChange = 0
            End If

            ' Print the results
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = yearlyChange
            'ws.Cells(summaryRow, 11).Value = percentageChange
            ws.Cells(summaryRow, 11).Value = Format(percentageChange, "0.00%") ' Format as percentage with 2 decimal places

            ws.Cells(summaryRow, 12).Value = totalVolume

' Apply conditional formatting for yearly change
                    If yearlyChange >= 0 Then
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 4 ' Green
                    Else
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 3 ' Red
                    End If

                    ' Apply conditional formatting for percentage change
                    If percentageChange >= 0 Then
                        ws.Cells(summaryRow, 11).Interior.ColorIndex = 4 ' Green
                    Else
                        ws.Cells(summaryRow, 11).Interior.ColorIndex = 3 ' Red
                    End If
            ' Increment the summary row
            summaryRow = summaryRow + 1
        End If
    Next i

    ' Autofit columns
    ws.Columns("I:L").AutoFit

End Sub


Results with all of AAB and subset of other tickers data when ran on 2018 sheet:
 

2.	Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
StockData2.vba
Sub StockData2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("2018") ' Change the sheet name as needed

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String

    greatestIncrease = -1
    greatestDecrease = 1
    greatestVolume = 0

    For i = 2 To lastRow
        Dim percentageChange As Double
        Dim totalVolume As Double

        percentageChange = ws.Cells(i, 3).Value
        totalVolume = ws.Cells(i, 4).Value

        If percentageChange > greatestIncrease Then
            greatestIncrease = percentageChange
            increaseTicker = ws.Cells(i, 1).Value
        End If

        If percentageChange < greatestDecrease Then
            greatestDecrease = percentageChange
            decreaseTicker = ws.Cells(i, 1).Value
        End If

        If totalVolume > greatestVolume Then
            greatestVolume = totalVolume
            volumeTicker = ws.Cells(i, 1).Value
        End If
    Next i

    ' Write results to new columns
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    

    ws.Cells(2, 16).Value = increaseTicker
    ws.Cells(2, 17).Value = Format(greatestIncrease * 100, "0.00%")
    ws.Cells(3, 16).Value = decreaseTicker
    ws.Cells(3, 17).Value = Format(greatestDecrease * 100, "0.00%")
    ws.Cells(4, 16).Value = volumeTicker
    ws.Cells(4, 17).Value = Format(greatestVolume, "#,##0")
End Sub
3.	Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
StockData3FinalCode that does all the calculations and new columns and in all sheets.
Sub StockData3FinalCode()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim i As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Instructions" Then ' Skip the Instructions sheet
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            summaryRow = 2
            greatestIncrease = -1
            greatestDecrease = 1
            greatestVolume = 0

            ' Set headers for the output columns
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

            For i = 2 To lastRow
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ' New ticker symbol
                    ticker = ws.Cells(i, 1).Value
                    openingPrice = ws.Cells(i, 3).Value
                    totalVolume = 0
                End If

                ' Add the volume to the total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ' Last row for the ticker symbol
                    closingPrice = ws.Cells(i, 6).Value
                    yearlyChange = closingPrice - openingPrice
                    If openingPrice <> 0 Then
                        percentageChange = yearlyChange / openingPrice
                    Else
                        percentageChange = 0
                    End If

                    ' Print the results
                    ws.Cells(summaryRow, 9).Value = ticker
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    ws.Cells(summaryRow, 11).Value = Format(percentageChange, "0.00%")
                    ws.Cells(summaryRow, 12).Value = totalVolume

                    ' Apply conditional formatting for yearly change
                    If yearlyChange >= 0 Then
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 4 ' Green
                    Else
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 3 ' Red
                    End If

                    ' Apply conditional formatting for percentage change
                    If percentageChange >= 0 Then
                        ws.Cells(summaryRow, 11).Interior.ColorIndex = 4 ' Green
                    Else
                        ws.Cells(summaryRow, 11).Interior.ColorIndex = 3 ' Red
                    End If

                    ' Increment the summary row
                    summaryRow = summaryRow + 1
                End If

                ' Find greatest % increase, % decrease, and total volume
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    increaseTicker = ticker
                End If

                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    decreaseTicker = ticker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTicker = ticker
                End If
            Next i

            ' Write results to new columns
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"

            ws.Cells(2, 16).Value = increaseTicker
            ws.Cells(2, 17).Value = Format(greatestIncrease * 100, "0.00%")
            ws.Cells(3, 16).Value = decreaseTicker
            ws.Cells(3, 17).Value = Format(greatestDecrease * 100, "0.00%")
            ws.Cells(4, 16).Value = volumeTicker
            ws.Cells(4, 17).Value = Format(greatestVolume, "#,##0")
        End If
    Next ws

    ' Autofit columns for all sheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit
    Next ws

End Sub


Code and results screenshot:
 



•	Solution File 
The dataset/solution file for this assignment can be found here. 
Multiple_year_stock_data.xlsm - https://github.com/pandarik/VBA-challenge/blob/main/Multiple_year_stock_data.xlsm
Code and results screenshots: https://github.com/pandarik/VBA-challenge/blob/main/ReadmeModule2ScreenshotsCode.docx


References/ Acknowledgements: 
Leveraged Google, ChatGPT, Copoilet as/where needed to develop/validate/troubleshoot code/data/functions.

Solution, Recommendations, and Conclusion:
In this VBA assignment, I leveraged VBA scripting to analyze generated stock market data. The objective was to create a script that could loop through all stocks for one year, calculating and outputting various key metrics such as yearly change, percentage change, and total stock volume. Additionally, the VBA script is designed to identify the stocks with the greatest increase, greatest decrease, and greatest total volume.

By completing this assignment, I think, I have demonstrated proficiency in using VBA scripting for data analysis in Excel. I have gained valuable experience in manipulating and analyzing large datasets, as well as in using conditional formatting to highlight specific data points. This assignment has furthered my skills as programmer and Excel expert, bringing me closer to achieving my goal of mastering the tools for data analysis and automation. I really liked the content, lecture, teachings, support, and assignments.

