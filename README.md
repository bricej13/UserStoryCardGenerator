# User Story Card Generator
This PowerShell script will run a Query from TFS and spit out the results into HTML cards that can be printed and used to track progress.

## Installation
1. Install Team Foundation Power Tools (http://www.microsoft.com/en-us/download/details.aspx?id=35775)
2. Change settings in file (I use PowerShell ISE)
  - **$usid**: is a single user story id. Set this variable if you only want the tasks from a single user story.
  - **$stretchstackrank**: is the threshold above which user stories will be considered 'stretch'. They print with a slightly different format for easy differentiation.
  - **$query**: is the path in TFS for the query you want to run. Note that it includes the project name, the folder structure, and the name of the query
  - **$collection**: is collection URI. This can be seen in the Team->Connect to Team Foundation Server window
3. Run the script
4. Print (Only Chrome has been confirmed to print properly)