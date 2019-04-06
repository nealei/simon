# [Simon](https://github.com/nealei/simon) <a href="https://github.com/nealei/simon/releases/latest"><img align="right" width="210" height="32" src="https://image.ibb.co/eRREgd/button_download_latest_release.png" alt="button_download_latest_release" border="0"></a><br />
<img align="right" width="300" height="300" src="https://proofpatisserie.files.wordpress.com/2012/06/img_2200.jpg">
**About:** Simon is a command-line interface (CLI) for the Trarr traffic simulation model, as an alternative to the GUI, which automates making and running Trarr traffic simulations. It is part of an Excel-Bat-Pie system (Bat as in Batch file and Pie as in Python - two scripting languages) and uses a simple approach to run simulations in about a quarter of the time of the Trarr Shell without tying up the user’s computer. The Simon batch file generates another batch file with the commands to run the Trarr simulation on each file.

**Get It:** The challenges of installing python in a corporate environment are reduced by installing Miniconda Python 3.7 without adding Anaconda to the PATH statement or registering Anaconda as the default Python. Start, Anaconda Prompt will ensure the correct environment is in place for the session. The make_trf.py and avg_seeds.py modules require the python openpyxl package and dependencies.
The computer must have the Trarr Traffic Simulation installed.
Install Simon files in the Trarr program directory. Save spreadsheet templates to automate the setup and reporting in the users template directory (eg C:\Users\user\AppData\Roaming\Microsoft\Templates).

**Use:** DoSimon and spreadsheet templates have been setup as an easy way of making and running simulations. Huge numbers of simulations can be set up in a spreadsheet and saved as a batch file which uses option_editor.py to make Trarr road option files without going through the graphical front-end or corrupting the files.

**1)** The proj_ directory must contain the existing road files (.ROD, .MLT & .OBS) for all time periods and years.

**2)** Traffic files can be created with the **TrarrSetupTRF** template.

**3)** The **MakeSimonSetup** template has instructions for saving a MakeSimonSetup file. Following is an extract from MakeSimonSetupHW4.bat:
```
Echo Run as %0 ^>%0.log to overwrite log file
..\option_editor.py HW4S3a.ROD 6.1 7.4 C HW4S3a_OT002.ROD "Overtaking Lane 2: 6.1-7.4km Eastbound"
..\option_editor.py HW4S3b.ROD 33 34 P HW4S3b_OT011.ROD "Overtaking Lane 11: 33.0-34.0km Westbound"
```
Note Limits: Overtaking lane start/stop one decimal place and Description 50 characters.   
This file is also used to setup economic analysis so you should familiarise yourself with file naming conventions in economic analysis template ReadMe sheet.

**4)** In Windows explorer navigate to proj_ directory and enter `..\mc` in Address Bar to start MiniConda prompt.

**5)** Enter `..\DoSimon`  for a simple run with all road and traffic files. Parameters with filename wildcards can be used to restrict the file set in the simulation. Eg:
```
..\DoSimon HW4_Sec3b* HW4S3b* HW4S3b > DoSimonSetupHW4S3b.log
```
(Process could be customised to append multiple Simon commands to single SimonSays.bat file for a road.)

**6)** If you have access to the **T17_TRARR economic analysis** template it can produce summary reports for all options in a road/traffic section.   
   Excel, File, New, My templates, ... follow instructions in ReadMe sheet.  
   Save to proj_ directory as type  ‘Excel Macro-Enabled Workbook (*.xlsm)'

**Support:** The system is modular and flexible. Be aware - you might find a use case that it will not work with so validate your results. Use the GitHub site to raise issues or contribute.

© Neale Irons, Licence: [CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/)
