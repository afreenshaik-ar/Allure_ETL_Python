# Allure_ETL_Python
Automated ETL testing using Python with database comparison and Allure reporting for clear, interactive test results.
This project is an ETL pipeline with test cases to compare data between source and target using Python and pytest. The tests include checks for table existence, record count, and data comparison, and generate an Allure report.

Test Cases
TC1 - Tables Exist: Check if source and target tables exist.

TC2 - Record Count Check: Validate if record counts match between the two queries.

TC3 - Data Compare: Compare data between two queries for accuracy.

Setup and Run Tests
Install dependencies:
pytest
selenium
allure-pytest


Run the tests with pytest: use below cmd
pytest -v -s --alluredir="<reports path>" testdb_compare.py


View Allure Report
After tests run, generate the Allure report:
Run following command in command prompt
allure serve <reports path>
