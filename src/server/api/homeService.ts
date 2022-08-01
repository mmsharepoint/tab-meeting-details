import { MicrosoftAppCredentials } from "botframework-connector";
const BotActivityHandler = require('../bots/BotActivityHandler');
// import BotActivityHandler from "../bots/BotActivityHandler";
import express = require("express");
import Axios from "axios";
import * as debug from "debug";
const log = debug("msteams");
const store = require('../api/store');

export const homeService = (options: any): express.Router => {
    const router = express.Router();

    async function getMeetingDetails(req, res)
    {
        const meetingId = req.params.meetingID;
        const credentials = new MicrosoftAppCredentials(process.env.MICROSOFT_APP_ID!, process.env.MICROSOFT_APP_PASSWORD!);
        const token = await credentials.getToken();
        const serviceUrl = store.getItem("serviceUrl");
        const apiUrl = `${serviceUrl}/v1/meetings/${meetingId}`;
        log(token);        
        const meetingDetails = store.getItem(`meetingDetails_${meetingId}`);
        log(meetingDetails);
        Axios.get(apiUrl, {
            headers: {          
                Authorization: `Bearer ${token}`
            }})
            .then((response) => {
                res.send(response.data);
            }).catch(err => {
                log(err);
                return null;
            });        
    }

    async function getMeetingParticipants(req, res)
    {
        const meetingId = req.params.meetingID;
        const participantId = req.params.userID;
        const tenantId = req.params.tenantID;
        const credentials = new MicrosoftAppCredentials(process.env.MICROSOFT_APP_ID!, process.env.MICROSOFT_APP_PASSWORD!);
        const token = await credentials.getToken();
        const serviceUrl = store.getItem("serviceUrl");
        const apiUrl = `${serviceUrl}/v1/meetings/${meetingId}/participants/${participantId}?tenantId=${tenantId}`;
        log(apiUrl);
        Axios.get(apiUrl, {
            headers: {
                Authorization: `Bearer ${token}`
            }})
            .then((response) => {
                res.send(response.data);
            }).catch(err => {
                log(err);
                return null;
            });        
    }

    router.post(
        "/getDetails/:meetingID",
        async (req: any, res: express.Response, next: express.NextFunction) => {
            getMeetingDetails(req, res);
        });
    
    router.post(
        "/getParticipants/:meetingID/:userID/:tenantID",
        async (req: any, res: express.Response, next: express.NextFunction) => {
            getMeetingParticipants(req, res);
        });
    return router;
}
