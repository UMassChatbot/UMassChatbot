const axios = require('axios');

/**
 * Calls the endpoint with authorization bearer token.
 * @param {string} endpoint
 * @param {string} accessToken
 */
async function callApiGetFreeBusy(endpoint, accessToken) {

    const options = {
        headers: {
            Authorization: `Bearer ${accessToken}`,
            'Prefer': 'outlook.timezone = "Eastern Standard Time"',
            'Content-Type': 'application/json'
        }
    };

    console.log('request made to web API at: ' + new Date().toString());

    //this is hardcoded to get nicholas norman's free/busy schedule within the date range 11/30 - 12/7
    //availabilityViewInterval sets the time intervals we are looking for
    try {
        axios.post(endpoint, {    
            Schedules: ['nnorman-umassd@devexch.umassp.edu'],
            StartTime: {
                dateTime: '2022-11-30T09:00:00',
                timeZone: 'Eastern Standard Time'
            },
            EndTime: {
                dateTime: '2022-12-7T18:00:00',
                timeZone: 'Eastern Standard Time'
            },
            availabilityViewInterval: "60"
            
        }, options)
            .then(function (response) {
            console.log(response.data);
        })
            .catch(function (error) {
            console.log(error);
        });

    } catch (error) {
        console.log(error)
        return error;
    }
}

//findMeetingTimes call
//needs a delegated token, findMeetingTimes does not support application function
async function callApiFindMeetingTimes(endpoint, accessToken) {

    const options = {
        headers: {
            Authorization: `Bearer ${accessToken}`,
            'Prefer': 'outlook.timezone = "Eastern Standard Time"',
            'Content-Type': 'application/json'
        }
    };

    console.log('request made to web API at: ' + new Date().toString());

    try {
        axios.post(endpoint, {    
            attendees: [ 
                { 
                  type: 'required',  
                  emailAddress: { 
                    name: 'Nicholas Norman',
                    address: 'nnorman-umassd@devexch.umassp.edu' 
                  } 
                }
              ],  
              locationConstraint: { 
                isRequired: false,  
                suggestLocation: false,  
                locations: [ 
                  { 
                    resolveAvailability: false,
                    displayName: 'Dion 110' 
                  } 
                ] 
              },  
              timeConstraint: {
                activityDomain: 'work', 
                timeSlots: [ 
                  { 
                    start: { 
                      dateTime: '2022-11-30T09:00:00',  
                      timeZone: 'Eastern Standard Time' 
                    },  
                    end: { 
                      dateTime: '2022-12-14T17:00:00',  
                      timeZone: 'Eastern Standard Time' 
                    } 
                  } 
                ] 
              },  
              meetingDuration: 'PT1H'
              
        }, options)
            .then(function (response) {
            //logs the meeting time suggestion
            console.log(response.data.meetingTimeSuggestions[0]);
        })
            .catch(function (error) {
            console.log(error);
        });

    } catch (error) {
        console.log(error)
        return error;
    }
}


//used for creating outlook calendar meetings
async function callApiCreateMeeting(endpoint, accessToken) {

    const options = {
        headers: {
            Authorization: `Bearer ${accessToken}`,
            'Prefer': 'outlook.timezone = "Eastern Standard Time"',
            'Content-Type': 'application/json'
        }
    };

    try {
        axios.post(endpoint, {
            subject: "Meeting",
            body: {
              contentType: "HTML",
              content: "Capstone meeting"
            },
            start: {
                dateTime: "2022-12-12T08:00:00",
                timeZone: "Eastern Standard Time"
            },
            end: {
                dateTime: "2022-12-12T09:00:00",
                timeZone: "Eastern Standard Time"
            },
            location:{
                displayName:"Dion 110"
            },
            attendees: [
              {
                emailAddress: {
                  address:"nnorman-umassd@devexch.umassp.edu",
                  name: "Nicholas Norman"
                },
                type: "required"
              }
            ]
          }, options)
          .then(function (response) {
            //console.log(response.data);
            console.log("Meeting created for: ")
            console.log(response.data.start.dateTime);
            console.log("At location: ")
            console.log(response.data.location.displayName);
            
        })
            .catch(function (error) {
            console.log(error);
        });
        } catch (error) {
            console.log(error)
            return error;
        }
}

//gets an application access token
//somewhat unneccesary since we have another function ot do this
async function callApiSignInClient(endpoint) {

    const options = {
        headers: {
             'Content-Type': 'application/x-www-form-urlencoded'
        }
    };

    console.log('request made to web API at: ' + new Date().toString());

    try {
        axios.post(endpoint, {    
            grant_type: 'client_credentials',
            client_id: process.env.CLIENT_ID,
            client_secret: process.env.CLIENT_SECRET,
            resource: process.env.GRAPH_ENDPOINT,
        }, options)
            .then(function (response) {
            console.log(response);
        })
            .catch(function (error) {
            console.log(error);
        });

    } catch (error) {
        console.log(error)
        return error;
    }
}


//gets delegated access token
//needed to do a findMeetingTimes call
async function callApiSignInUser(endpoint) {

    const options = {
        headers: {
             'Content-Type': 'application/x-www-form-urlencoded'
        }
    };

    console.log('request made to web API at: ' + new Date().toString());


    //stop here
    try {
        const delegatedTokenRequest = axios.post(endpoint, {    
            grant_type: 'password',
            client_id: process.env.CLIENT_ID,
            client_secret: process.env.CLIENT_SECRET,
            resource: process.env.GRAPH_ENDPOINT,
            username: process.env.LOGIN,
            password: process.env.PASSWORD
        }, options);

        const delegatedToken = delegatedTokenRequest.then((response) => response.data.access_token);

        return delegatedToken;

    } catch (error) {
        console.log(error)
        return error;
    }
}

module.exports = {
    callApiGetFreeBusy: callApiGetFreeBusy,
    callApiFindMeetingTimes : callApiFindMeetingTimes,
    callApiCreateMeeting : callApiCreateMeeting,
    callApiSignInClient : callApiSignInClient,
    callApiSignInUser : callApiSignInUser,

};