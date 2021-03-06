# Canadian-tax-calculator
LibreOffice Calc integration to Wealthsimple 2021 Free Income Tax Calculator

The Tools->Macros->CalculateTaxes function in the Sample.ods spreadsheet file is embedded python code that integrates to the browser based Wealthsimple 2021 free income tax calculator. It quickly estimates your taxes, provides visibility of your tax bracket, marginal tax rates, and average tax rates. It repeats the tax calculation for all supported Free Income Tax Calculator input fields, once for each input row, and records the result into the corresponding output rows in the Model spreadsheet.

This software is meant to help understand wealth accumulation and decumulation strategies and quickly estimate related tax implications.

![Figure 1: CalculateTaxes](Documentation/MediaPreview.png?raw=True "Figure 1: CalculateTaxes")

This project was last tested on April 13th, 2022 and continues to be fully functional. 

![Figure 2: Run the CalculateTaxes Python Macro](Documentation/ToolMenuMacros.png?raw=True "Figure 2: Run the CalculateTaxes Python Macro")

CalculateTaxes macro delivers easily accessible data that anyone who tinkers with spreadsheets can use to make a fully customized application. Nobody ought to use software, including code from this repository, to access the internet without first being able to perform a security audit of the code.

**Prerequistes**

<ul>
   <li>[LibreOffice 7.2.5.2] (https://www.libreoffice.org/download/download/)</li>
   <li>[Python 3.8.10] (https://www.python.org/downloads/release/python-3810/)</li>
   <li>[ChromeDriver 100.0.4896.60] (https://chromedriver.storage.googleapis.com/index.html?path=100.0.4896.60/)</li>
   <li>[Chromium based browser Version 1.37.113 Chromium: 100.0.4896.88 (Official Build) (64-bit)] (https://brave.com/download/)</li>
   <li>[Python site-packages] certifi, selenium, urllib3 installed in C:\Program Files\LibreOffice\program\python-core-3.8.10\lib\site-packages</li>
</ul>

**How To Use**

The spreadsheet must have a sheet named "Model" and a sheet named "Configuration". The "Configuration" spreadsheet must have the "driver", "browser", and "url" fields properly configured. Only chromium-based browsers will work. It has been tested using Brave which is based on Chrome version 98. A chromedriver that matches the chromium version must be installed. Apparently any chromium based browser will work as expected. The Wealthsimple url is based on province.

The model must have the LibreOffice Calc named ranges listed below. A current limitation is that the Model must have a named range for otherIncome. Each row of the named input ranges are supplied to Wealthsimple 2021 Free Income Tax Calculator one row at a time and the corresponding results returned to the same row for the named output ranges.

The Wealthsimple 2021 Free Income Tax Calculator supports the following web page input fields that are mapped to a corresponding LibreOffice Calc named range:
* input field employmentIncome -> LibreOffice Calc named range IN_employmentIncome
* selfEmploymentIncome -> IN_selfEmploymentIncome
* rrspDeduction -> IN_rrspDeduction
* capitalGains -> IN_capitalGains
* eligibleDividends -> IN_eligibleDividends
* ineligibleDividends -> IN_ineligibleDividends
* otherIncome -> IN_otherIncome
* incomeTaxesPaid -> IN_incomeTaxesPaid

If the named input range does not exist in the Model spreadsheet file then a value of 0 is supplied to the Wealthsimple 2021 Free Income Tax Calculator for that field.

The Wealthsimple 2021 Free Income Tax Calculator supports the following web page output fields that are mapped to a corresponding LibreOffice Calc named range:
* totalIncome -> OUT_totalIncome
* totalTax -> OUT_totalTax
* afterTaxIncome -> OUT_afterTaxIncome
* averageTaxRate -> OUT_averageTaxRate
* marginalTaxRate -> OUT_marginalTaxRate

If the named output range does not exist in the Model spreadsheet file then the Wealthsimple 2021 Free Income Tax Calculator result is ignored for that field.

**User Notes**

LibreOffice must be installed to make use of the included Sample.ods spreadsheet file. Python and Basic macros are pre-installed into the spreadsheet by default so no additional programming is required to make the spreadsheet work. The file is a standard ODS archive generated by LibreOffice and can be fully customized using LibreOffice spreadsheet functions without additional software development.

The Sample.ods spreadsheet is structured to help understand decumulation strategies. It can be used to help understand situations like how to empty an RRSP before 70? What are the tax implications? How do I decumulate all investments and achieve a consistent inflation adjusted income for a lifetime, and then expire without a penny remaining? Values in the Sample spreadsheet have been floated in r/fican and I wanted to see what is required to get a reasonably complete answer ... or at least have a tool that might lead to good answers. The Sample is an excellent demonstration of Wealthsimple 2021 Free Income Tax Calculator integration but the user has to do the heavy lifting to understand results and how they might be used for accumulation and decumulation planning.

![Figure 3: A consistent inflation adjusted lifetime income, and then expire without a penny example](Documentation/Sample-SmoothRRIFWithdrawal.png?raw=True "Figure 3: A consistent inflation adjusted lifetime income, and then expire without a penny example")

**Canadian-tax-calculator is free software: you can redistribute and/or modify it under the terms of the MIT License.**
