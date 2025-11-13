# Calculate the number of qualifiers by age group and event
- add a folder called data, which contains your swimrankings data
- change the const TIME_STANDARDS_FILE: &str = "timestandards.xlsx"; to whatever your standards file is named
- XLSX files are currently only file types that are being handled
- program will output file named qualifiers_count.xlsx in the root directory of the program



# Features
- normalizes event names 
  - Bu -> Fly
  - ME -> IM
- finds best suited age group based on the standards file:
  - will calculate the &under categories, and the &over categories based on age groups for standards used
 

# Debug
- added console log to show column headings for standards file
- added console tog show files being parsed
- added console log to show total files contained in /data folder
- added console log to show total results, and number of qualifiers found


# Console Log Sample:
- Total results extracted: 16547

Sample results:
  Sex: Men, Age: 12, Event: 100Bk, Time: 73.37s
  Sex: Men, Age: 12, Event: 100Bk, Time: 73.43s
  Sex: Men, Age: 12, Event: 100Bk, Time: 74.01s

Ages found in meet data: ["11", "12", "13", "14", "15", "16", "18"]
Events found in meet data: ["50Bk", "4x50Fr", "50BuLap", "200Me", "50Bu"]

Counting qualifiers...
DEBUG: Found 116 qualifying times
DEBUG: 6168 results had no matching standard
Found 81 qualifier count entries

Analysis complete! Results saved to qualifier_counts.xlsx
