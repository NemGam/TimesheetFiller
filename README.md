# TimesheetFiller
This is a script to fill timesheets for the University of South Florida automatically.

## Requirements:
  - Windows OS
  - Python (>= 3.10.11)
  - pipreqs package
    ```bash
    pip install pipreqs
    ```
  - ICal link with your schedule
## Installation (Windows only):
  1. If you don't have the required packages (as listed in the requirements.txt file) or do not want to install them, you can download the *-onefile version of the script. (Skip to step 4)
  2. Clone the repository or download timesheet-filler.py and requirements.txt into some folder.
  3. Run the following command from the same folder:
    ```pip install -r requirements.txt```
  4. Put your empty timesheet with all pages into the same folder as the .py script.
  5. Run the script
    ```python3 timesheet-filler.py```
     (or ```python3 timesheet-filler-onefile.py```)
  6. If it is the first time, you'll be asked to input your full name, the ICal link, and the name of the event (the event name that depicts your time on the calendar).
     - You can change it later by going into the config.json file or deleting it.
  7. After some time, the filled Excel sheet and the .pdf file will be generated. **Make sure that it fills your time properly!**

## Feedback
If you have any additional ideas or feedback, feel free to send me a message at nebuglaz@gmail.com
