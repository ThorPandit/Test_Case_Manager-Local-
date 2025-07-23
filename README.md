# Test_Case_Manager-Local-
This used to manage the tracking of defined/undefined test cases for a DUT to be performed. This is based on local system using user name and password.  
Notes:-
Please make sure you login via admin to import the test case or to add test case as user can only test and make the report. 
You can attatche one file for each test case and the file name will be linked in the final report generation. 
Make sure the excel you are importing must follow below format:-
<img width="963" height="109" alt="image" src="https://github.com/user-attachments/assets/1e607f37-ed59-4f65-8b91-f008198e8f7b" />
Please amke sure you have installed all dependencies of python. 

To generate the exe please use below run command to execute:-
pyinstaller --onefile --noconsole --icon=youricon.ico --add-data "users.json;." --add-data "test_cases.json;." main.py

-icon=youricon.ico: if you are using icon as .ico format
users.json: make sure users.json is in same folder of main.py
test_cases.json: make sure test_cases.json is in same folder of main.py

login will look like this:-
<img width="195" height="99" alt="image" src="https://github.com/user-attachments/assets/b792fa05-49a2-4708-bb66-4e85451859dc" />

enter the credentials and click on login.
<img width="200" height="98" alt="image" src="https://github.com/user-attachments/assets/84dd1523-1f5c-4c45-9916-db2baa0377e8" />

if you loign via admin credentials it will open like-
<img width="1199" height="326" alt="image" src="https://github.com/user-attachments/assets/4aff37b0-8050-4437-9e70-716a94796382" />

with test manual additon it will look like:-
<img width="1196" height="324" alt="image" src="https://github.com/user-attachments/assets/c6ec3a27-32dc-410f-bdfe-7e14740b63d4" />

with excel it will import directly. so use as required. 

you can contact me if issue occur on: **thorpandit@gmail.com**
Make sure to make subject like: **"GITHUB Test case manager query from #your name" **
