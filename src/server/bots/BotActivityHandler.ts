import { BotDeclaration } from "express-msteams-host";
import { CardFactory, ConversationState, MemoryStorage, TaskModuleRequest, TaskModuleResponse, TaskModuleTaskInfo, TeamsActivityHandler, TeamsInfo, TurnContext, UserState } from "botbuilder";
import * as debug from "debug";
import { getMeetingDetailsCard } from "../../model/meetingDetailsCard";
import { IMeetingDetails } from "../../model/IMeetingDetails";
const log = debug("msteams");
const store = require('../api/store');

@BotDeclaration(
  "/api/messages",
  new MemoryStorage(),
  // eslint-disable-next-line no-undef
  process.env.MICROSOFT_APP_ID,
  // eslint-disable-next-line no-undef
  process.env.MICROSOFT_APP_PASSWORD)
export class BotActivityHandler extends TeamsActivityHandler {
  constructor(public conversationState: ConversationState, userState: UserState) {
    super();
    this.conversationState = conversationState;
    this.onMessage(async (context, next) => {        
        await context.sendActivity("Welcome to Meeting Details Application!");            
    });

    this.onConversationUpdate(async (context, next) => {
      store.setItem("serviceUrl", context.activity.serviceUrl);
    });
  }

  protected async handleTeamsTaskModuleSubmit(_context: TurnContext, _taskModuleRequest: TaskModuleRequest): Promise<any> {
    log(_context.activity);    

    switch (_taskModuleRequest.data.verb) {
      case "getMeetingDetails":
        const meetingID = _taskModuleRequest.data.data.meetingId;
        const meetingDetails = await TeamsInfo.getMeetingInfo(_context, meetingID) as IMeetingDetails;

        const card = getMeetingDetailsCard(meetingDetails);

        const Response: TaskModuleResponse = {
          task: {
            type: 'continue',
            value: {
              title: "Your Meeting Details",
              height: 500,
              width: "large",
              card: CardFactory.adaptiveCard(card),
            } as TaskModuleTaskInfo
          }
        };
        return Promise.resolve(Response);
      default:
        store.setItem("serviceUrl", _context.activity.serviceUrl);
        return null;
    }
}
}
