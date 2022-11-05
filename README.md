# Data Ingestion With Selenium V1
## _Automatically filling multiple responses into a **single page** Google Form with Selenium_ and Python
![Cover](https://github.com/RealXun/Selenium-Ingestion-Python/blob/main/Resources/cover.png)
> Note: Check below for a video example.
## Prerequisites:
- Selenium
- Python
- Google Chrome
- ChromeDriver
- Microsoft Excel

The Task here is to fill multiple responses from a excel file with the same Google form using selenium in python. Link to the Google forms and excel file used in this example are given below:

## Links:
- InformationCalls Form – [Click Here](https://forms.gle/jm28YptQGPj6XvLJA)
- Excel File – [Click Here](https://github.com/RealXun/Selenium-Ingestion-Python/blob/main/Calls.xlsx)
- Python File– [Click Here](https://github.com/RealXun/Selenium-Ingestion-Python/blob/main/Data_Ingestion_Single_Page.py)

## The form has six entries:

- Agent Name
   - Daniel
   - David
- Source (Source of the call)
   - Incoming Call
   - Outgoing Call
- Phone Number
- Email Adress
- Product
  - Home repairs
  - Personal care
- Call Result
  - Silent call
  - Call dropped
  - Unanswered call
  - Transferred to sales

![form](https://github.com/RealXun/Selenium-Ingestion-Python/blob/main/Resources/Form%20Picture.PNG)

## Excel File:
The exel file should be modified as needed. Always according to the changes in the code.

> Note: This excel is ready for two types of forms. But this code only use the information form. Also you will need to adjust the route where you have the excel file.

In the image we can see what rows we use in this form

![Excel](https://github.com/RealXun/Selenium-Ingestion-Python/blob/main/Resources/Information%20Excel%20Image.PNG)

## Approach:
Using the Xpath of the google form we want to achieve our required functionality given steps needs to followed in an order:

- Import selenium and time module
- Excel File Path and Read Excel File
- Add chrome driver path
- Create dicctionaries
- Creation of Variables
- Open chrome and load google form
- Iterate through each data and fill detail
- Log File Creation
- Close the window

## Log File:
A log file will be created in the same folder where you execute the code to check if everything was ok or to check where the code failed ingesting data.

## Try the code:
If you try this code with the same excel file, you can check the ingested data in [this link](https://docs.google.com/spreadsheets/d/1MEjJ1B7DBeAALJrZeE3LAuyaFqDc835yOe_sSm8RNdg/edit?usp=sharing)

![ExcelOnline](https://github.com/RealXun/Selenium-Ingestion-Python/blob/main/Resources/Excel%20online%20image.PNG)

## video example:
https://user-images.githubusercontent.com/112812176/190384467-a29c3ee7-7f59-466f-85e2-cf6b6bb28728.mp4

## Contact:
Do not hesitate to contact me for any further explanation about the code.

## License
MIT

**Free Software, Hell Yeah!**
