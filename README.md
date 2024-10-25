# Automated REM Sleep without Atonia Detection Algorithm (AURA)
-Provides RWA quantification results by CRWA%, SINBAR, AASM, and RAI methods.

## Features
- Load and analyze REM sleep data from `.edf` and `.xlsx` files.
- Export analysis results folder.
- Results Files: comb_CRWA, comb_SINBAR, comb_AASM, comb_RAI

## usage
1. Run the application:
2. Use the "Select Data File" button to load an `.edf` file.
3. Use the "Select Event File" button to load an `.xlsx` event file from RemLogic.
4. Use the "Select Result Folder" button to choose where to save the results.
5. Click "Start Analysis" to begin the process.

## event file for using
1. In REMLogic, File -> Export -> Events -> Add All
2. Options, Check only show position column and Sleep stage column
3. Convert txt file to EXCEL file
4. Check the export file if there is error in time format!!

## License
This project is licensed under the MIT License.

## Contributors
- Tae-Gon Noh (Main Deveopler)
- Contact if you find error: ghgh3110@snu.ac.kr

## Acknowlegements
- Thanks to Professor Ki-Yong Jung for guidance.
- Special thanks to Dr. Shin and Lee for validation.
