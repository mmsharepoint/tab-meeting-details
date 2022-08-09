export const TaskCard = {
  type: "AdaptiveCard",
  body: [
    {
      type: "TextBlock",
      size: "Medium",
      weight: "Bolder",
      text: "Meeting Details"
    },
    {
      type: "TextBlock",
      text: "This actions sends current meeting Details to the chat.",
      wrap: true
    }
  ],
  actions: [{
    type: "Action.Submit",
    id: "SndDetails",
    title: "Send Details",
    data: {
      msteams: {
        "type": "task/submit"
      },
      verb: "getMeetingDetails",
      data: {
        meetingId: ""
      }
    }
  }],
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  version: "1.4"
};
  