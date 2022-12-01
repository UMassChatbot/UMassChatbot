This application is built off of the "Tutorial: Call the Microsoft Graph API in a Node.js console app" 
on Microsoft's Graph API Documentation, founde here: 
https://learn.microsoft.com/en-us/azure/active-directory/develop/tutorial-v2-nodejs-console

Authentication goes through the 'NodeConsoleApp' on the Azure Active Directory.

It has been modified to alter the API calls it can make, namely the getSchedule and findMeetinTimes calls.
It is hardcoded to look get the schedule of and find meeting times with our team lead,
Nicholas Norman, and the logged in user (as long as they are within the UMass Pres Office system).

The application can be run by inputting "node . --op [mode]" into the terminal.
Current valid modes are: "getFreeBusy" and "findMeetingTimes", e.g. "node. --op getFreeBusy".

The attached body is within the fetch.js file and contains the free/busy body currently, 
needs to be commented out and the other body commented in to do a findMeetingTimes request.

We currently receive 400, 401, and 403 errors depending on changes we have made but are unsure what is invalid.

Relevant documention:
https://learn.microsoft.com/en-us/graph/outlook-get-free-busy-schedule
https://learn.microsoft.com/en-us/graph/findmeetingtimes-example
https://axios-http.com/docs/post_example