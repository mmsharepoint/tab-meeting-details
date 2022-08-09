import * as React from "react";
import { Provider, Flex, Text, Button, Header, Menu, tabListBehavior } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, TaskInfo, tasks } from "@microsoft/teams-js";
import Axios from "axios";
import { IMeetingDetails } from "../../model/IMeetingDetails";
import { IMeetingParticipant } from "../../model/IMeetingParticipant";
import { TaskCard } from "../../model/taskCard";
import { MeetingDetails } from "./components/MeetingDetails";
import { MeetingParticipant } from "./components/MeetingParticipant";

/**
 * Implementation of the Meeting Details content page
 */
export const MeetingDetailsTab = () => {
  const [{ inTeams, theme, context }] = useTeams();
  const [entityId, setEntityId] = useState<string | undefined>();
  const [meetingDetails, setMeetingDetails] = useState<IMeetingDetails>();
  const [meetingParticipant, setMeetingParticipant] = useState<IMeetingParticipant>();
  const [activeMenuIndex, setActiveMenuIndex] = useState<number>(0);

  const menuItems = [
    {
      key: 'meetingDetails',
      content: 'Meeting Details',
    },
    {
      key: 'meetingParticipant',
      content: 'Meeting Participant',
    }
  ];

  const getDetails = async (meetingID: string) => {        
    const response = await Axios.post(`https://${process.env.PUBLIC_HOSTNAME}/api/getDetails/${meetingID}`);
    setMeetingDetails(response.data); 
  };

  const reloadDetails = () => {
    TaskCard!.actions![0].data!.data.meetingId = context?.meeting!.id!;
    const taskCardAttachment = {
                                contentType: "application/vnd.microsoft.card.adaptive",
                                content: TaskCard } 
    const taskModuleInfo: TaskInfo = {
        title: "Snd Details",
        card: JSON.stringify(taskCardAttachment),
        width: 300,
        height: 250,
        completionBotId: process.env.MICROSOFT_APP_ID
    };

    tasks.startTask(taskModuleInfo, reloadDetailsCB);
  };

  const reloadDetailsCB = () => {
      // tasks.submitTask({ meetingId: context?.meeting?.id }, process.env.MICROSOFT_APP_ID);        
      tasks.submitTask();
  };
  const getParticipant = async (meetingID: string, userId, tenantId) => {        
    const response = await Axios.post(`https://${process.env.PUBLIC_HOSTNAME}/api/getParticipantDetails/${meetingID}/${userId}/${tenantId}`);
    setMeetingParticipant(response.data);
  };

  const onActiveIndexChange = (event, data) => {
    setActiveMenuIndex(data.activeIndex);
  };

  useEffect(() => {
    if (inTeams === true) {
      app.notifySuccess();
    } else {
      setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams]);

  useEffect(() => {
    if (context) {
      setEntityId(context.page.id);
      if (context.meeting) {
        getDetails(context.meeting.id);
        getParticipant(context.meeting.id, context.user?.id, context.user?.tenant?.id);
      }
    }
  }, [context]);

  /**
   * The render() method to create the UI of the tab
   */
  return (
    <Provider theme={theme}>
      <Flex fill={true} column styles={{
          padding: ".8rem 0 .8rem .5rem"
      }}>
        <Flex.Item>
          <Header content="Meeting Information" />
        </Flex.Item>
        <Flex.Item>
          <div>
            <Menu
                defaultActiveIndex={0}
                activeIndex={activeMenuIndex}
                onActiveIndexChange={onActiveIndexChange}
                items={menuItems}
                underlined
                primary
                accessibility={tabListBehavior}
                aria-label="Meeting Information"
            />
            <div className="l-content">
                {activeMenuIndex === 0 && <MeetingDetails meetingDetails={meetingDetails} reloadDetails={reloadDetails} />}
                {activeMenuIndex === 1 && <MeetingParticipant meetingParticipant={meetingParticipant} />}
            </div>
          </div>
        </Flex.Item>
      </Flex>
    </Provider>
  );
};
