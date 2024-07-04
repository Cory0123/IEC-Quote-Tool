How to Run the Script ... by Ted

Step 1: Open Terminal in Visual Studio Code
    Click on the "Terminal" tab in the top-left corner.
    Click on "New Terminal".
Step 2: Run the Script to generate "Summary_AMO.xlsx"
    In the terminal, type the following command:
        python AMO.py
Step 3: Run the Script to copy some data from "Inventec CQ OPP & cNB & Docking July 2024 Initial Quote - Copy.xlsx" to "Summary_AMO.xlsx" sheet "data1"
    In the terminal, type the following command:
        python AMO1.py      
Step 4: Run the Script to do merge(vLookUp) between sheets "Sheet1", "data1" and generate to "merged_data2" in "Summary_AMO.xlsx"
    In the terminal, type the following command:
        python AMO2.py      
