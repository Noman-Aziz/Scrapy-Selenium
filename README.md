# Scrapy-Selenium
This is a Python Scrapy Selenium Project for the [Website](https://www.heiminfo.ch/institutionen).

Functionalities:
* LinkExtractor.py File is Used to Extract All the Urls from that Webpage using Selenium and add them to Text File.
* First We have to Click on Some Buttons on the Website to generate Results, which is handled through Selenium headless browser.
* The Website Uses JavaScript and Show-more Button, so Pagination is handled through Selenium headless browser.
* Then The Scrapy Spider Reads the Text File and Add all The Urls in **Start_Urls** List.
* Then The Scrapy Spider Reads The Fields to Extract from that Website using a *Parameter.xlsx* file using *Openpyxl* module.

Required Python Modules:
* Scrapy
* Selenium
* Scrapy_xlsx
* Scrapy_user_agents
* Openpyxl
