import { BotDeclaration } from "express-msteams-host";
import { ConversationState, MemoryStorage, TeamsActivityHandler, TeamsInfo, UserState } from "botbuilder";
import * as debug from "debug";
const log = debug("msteams");
const store = require('../api/store');

let ConversationID = "";
let serviceUrl = "";

@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_ID,
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_PASSWORD)
class BotActivityHandler extends TeamsActivityHandler {
    // public serviceUrl = "";

    constructor(public conversationState: ConversationState, userState: UserState) {
        super();
        this.conversationState = conversationState;
        this.onMessage(async (context, next) => {        
            await context.sendActivity("Welcome to SidePanel Application!");            
        });

        this.onConversationUpdate(async (context, next) => {
            ConversationID = context.activity.conversation.id;
            exports.ConversationID = ConversationID;
            serviceUrl = context.activity.serviceUrl;
            // exports.serviceUrl = serviceUrl;
            store.setItem("serviceUrl", serviceUrl);
            try {
                const meetingID = context.activity.channelData.meeting.id; // "MCMxOTptZWV0aW5nX01tRXpNRGhsTmpRdE1HVXdNQzAwWlRGakxUZ3lOMkV0TnpnME5XTmhNRFE1Tm1NeEB0aHJlYWQudjIjMA==";
                const meetingDetails = await TeamsInfo.getMeetingInfo(context, meetingID);
                log(meetingDetails);
                store.setItem(`meetingDetails_${meetingID}`, meetingDetails);
            }
            catch(err) {
                log(err);
            };
            
        });

        
        // this.onMembersAddedActivity(async (context, next) => {
        //     context.activity.membersAdded.forEach(async (teamMember) => {
        //         if (teamMember.id !== context.activity.recipient.id) {
        //             await context.sendActivity(`Welcome to the team ${ teamMember.givenName } ${ teamMember.surname }`);
        //         }
        //     });
        //     await next();
        // });
    }
}

module.exports.BotActivityHandler = BotActivityHandler;