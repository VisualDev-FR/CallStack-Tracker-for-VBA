# CallStack-Tracker-for-VBA
A single VBA module wich sequentially write the call stack in html file during the runtime. It allows you to quickly find execution errors, and / or anormal function leaves.

# Installing
1) Add the 'Microsoft-sripting-runtime' to the VBA references (VBA IDE / tools / References / Microsoft-sripting-runtime)
2) Initialize your file system
  - Import the .bas file into your VBA Project
  - Insert your VBA Project (.xlsm, .xlam) into the root directory containing the directories HTML / CRASH / REPORT
  - You can adapt your fileSystem organisation by modifying the function 'InitFileSystem'

# Track your callStack
1) In each sub / function you want to track, insert the following line :
  - Call XX_PRINT_HISTO_XX('MODULE_NAME', 'SUB_NAME') -> MODULE_NAME correspond to the cuurent Module/Class and SUB_NAME to the current sub/function you want to track

2) At the normal end of your sub, insert the following line :
  - Call XX_PRINT_HISTO_XX(histEND) -> the tracker will detect the normal end of your sub
 
3) At the anormal leaving procedure handling, insert the following line :
  - Call XX_PRINT_HISTO_XX(histEXIT) -> the tracker will highlight the anormal sub leaving and will show it to you in the global report

4) In yout error handling, insert the following line :
  - Call XX_PRINT_HISTO_XX(histERROR) -> the tracker will highlight the error, and will show it to you in the global report with the error description

5) Open your CallStack Report in any Web browser, see what happened during the runtime, and analyse your execution times step by step.
  - You can find all your reports in the directory 'REPORT'
  - All reports with at least one error/anormal exit will be sent in the directory 'CRASH'