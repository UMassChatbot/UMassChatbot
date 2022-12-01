#!/usr/bin/env node

// read in env settings
require('dotenv').config();

const yargs = require('yargs');

//import variables from other files
const fetch = require('./fetch');
const auth = require('./auth');
const { AppMetadataEntity } = require('@azure/msal-common');

//type 'node . --op [operation]' in terminal to make a request
const options = yargs
    .usage('Usage: --op <operation_name>')
    .option('op', { alias: 'operation', describe: 'operation name', type: 'string', demandOption: true })
    .argv;

async function main() {
    console.log(`You have selected: ${options.op}`);

    switch (yargs.argv['op']) {

        //this call should find meeting times between nicholas norman and the user specified in the endpoint
        case 'findMeetingTimes':
            try {

                //the correct token that we need is a delegated token
                //we can get this token by making a callApiSignInUser call
                //and pull the access_token out of the response

                //make a delegated token request
                //returns a delegated access token
                const delegatedToken = await fetch.callApiSignInUser(auth.apiConfigSignInUser.uri);

                //may need an additional await/then
                
                //token not getting passed correctly, but works when hard coded
                //make a signInUser call to get a delegated token and then replace delegatedToken with the actual token(in '')

                //make the findMeetingTimes call with the delegated token
                const meetingTimes = await fetch.callApiFindMeetingTimes(auth.apiConfigMeetingTimes.uri, delegatedToken);

                // display result
                console.log(meetingTimes);
            } catch (error) {
                console.log(error);
            }
            
        break;

        //this call should return nicholas norman's free/busy schedule
        case 'getFreeBusy' :
            try {
                const authResponse = await auth.getToken(auth.tokenRequest);
                const freeBusySchedule = await fetch.callApiGetFreeBusy(auth.apiConfigGetFreeBusy.uri, authResponse.accessToken);
                console.log(freeBusySchedule);
            } catch (error) {
                console.log(error);
            }

        break;

        //create a calendar event
        case 'createMeeting' :
            try {
                const authResponse = await auth.getToken(auth.tokenRequest);
                const meeting = await fetch.callApiCreateMeeting(auth.apiConfigCreateMeeting.uri, authResponse.accessToken);
                console.log(meeting);
            } catch (error) {
                console.log(error);
            }

        //get an application access token
        case 'signInClient' :
            try {
                const signedInClient = await fetch.callApiSignInClient(auth.apiConfigSignInClient.uri);
                console.log(signedInClient);
            } catch (error) {
                console.log(error);
            }

        break;

        //get a delegated access token
        //necessary for some requests
        case 'signInUser' :
            try {
                const signedInUser = await fetch.callApiSignInUser(auth.apiConfigSignInUser.uri);
                console.log(signedInUser);
            } catch (error) {
                console.log(error);
            }

        break;

        default:
            console.log('Select a Graph operation first');
        break;
    }
    
};

main();