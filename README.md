# How to use the script:

# Step 1. MAKE A COPY OF YOUR GRID 
   - This is just in case something gets messed up for whatever reason.
# Step 2. Download Zoom attendance reports
   - Download all of the Zoom reports that you want to record in your grid.
   - The reports can be reached by clicking Reports->Usage
   - Make sure you check "Show unique users" before pressing export. This is important.
   - Don't worry about your name showing up in the report, the script accounts for that.
   - As you download each one, make sure that you are naming it after the day of the month that the meeting happened.
   - For example, if you are downloading the attendance report for 2/18/21, name it 18
   - Here is an image of what your folder should look like once you have downloaded several attendance reports:
![IMAGE](Pictures/folder.png)
# Step 3. Open the grid
   - You will be asked a few questions about your grid since each grid has a different number of students.
   - Make note of the number of the first blank row in your grid.
   - Look at the image below as an example. The number of the first blank row is 506. If the 2 names were not there then it would be 504.
![IMAGE](Pictures/grid1.png)
   - The next row number to make a note of is the row where all of the attendance gets added up on. We would not want to overwrite that.
   - In the image below, the row number we would want is 554. Make sure you don't get it confused with 503, which also has the sums of attendance.
![IMAGE](Pictures/grid2.png)
# Step 4. Close the grid and all attendance reports
   - If these are open when we run the script then it will not work.
# Step 5. Make sure that you have the following in the same folder:
   - Grid
   - Script
   - All attendance reports
# Step 6. Download the following program:

   - Python3 (Make sure you download python 3 instead of python 2)

     - Download for windows: https://www.howtogeek.com/197947/how-to-install-python-on-windows/ 
       - (stop when you reach Adjust System Variables So You Can Access ...)

       - open command prompt and type (without quotes) "pip install openpyxl"

     - for mac(untested): https://installpython3.com/mac/ (stop when you reach the Virtual Environments heading)
       - open the terminal and type (without quotes) "pip install openpyxl"

# Step 7. Run the script
   - using the terminal(mac) or command prompt(windows) navigate to the folder where you have all of the things from step 5.
   - This can be done by going from folder to folder using cd 'folder name'
   - If you are unsure how to do this then definitely watch the video linked at the top which goes over all of these steps in greater detail.
   - Once you are on the correct folder in the terminal, type the following without quotes "python script.py"
   - Answer all of the questions properly. I do recommend looking at the video just to make sure you input the right things.
