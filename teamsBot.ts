import { TeamsActivityHandler, TurnContext, MessageFactory, ActivityTypes, CardFactory, AdaptiveCardInvokeResponse, AdaptiveCardInvokeValue } from "botbuilder";
import { getChatSummary, createPDF, uploadPDFToSharePoint } from './service';
// import { openAdaptiveCard } from "./adaptivecard";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";

// PS_02
export class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();

        // PS_06
        this.onMessage(async (context, next) => {
            console.log("Running with Message Activity.");

            if (context.activity.text) {
                const removedMentionText = TurnContext.removeRecipientMention(context.activity);
                const txt = removedMentionText ? removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim() : "";
                const command = context.activity.text;

                // PS_06
                if (command.toLowerCase().includes("chat summarise") || command.toLowerCase().includes("summary") || command.toLowerCase().includes("summarise")) {
                    await context.sendActivity({ type: ActivityTypes.Typing });
                    await this.summarizeChatHistory(context);
                } else {
                    await context.sendActivity({ type: ActivityTypes.Typing });
                    await context.sendActivity(`You said: "${command}"`);
                }
            } else {
                await context.sendActivity(`No text detected in the message.`);
            }

            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id) {
                    await context.sendActivity(
                        `Hi there! I'm a Teams bot, I can Give you the chat summary`,
                    );
                    break;
                }
            }
            await next();
        });
    }
    async onAdaptiveCardInvoke(context: TurnContext, invokeValue: any): Promise<any> {

        return { statusCode: 200 };
    }

    // PS_05
    async summarizeChatHistory(context: TurnContext): Promise<void> {
        // PS_07
        const chatMessages = await this.fetchChatHistory(context);
        // PS_18
        const chatText = chatMessages.map(message => message.body.content).join("\n");
        // PS_19
        const summary = await getChatSummary(chatText);
        // PS_22
        await this.postChatSummary(context, summary);
        // PS_25
        await this.uploadAndSendPDF(context, summary);
    }

    // PS_08 - PS_17
    async fetchChatHistory(context: TurnContext): Promise<any[]> {
        debugger
        const tenantId = process.env.AZURE_TENANT_ID || 'cd5dbd24-330a-4174-8e00-dc550f63976b';
        const clientId = process.env.AZURE_CLIENT_ID || 'e1cd3b9a-0ae7-467c-847f-07f521343b85';
        const clientSecret = process.env.AZURE_CLIENT_SECRET || '5G.8Q~O0yZ3PSHoFTwF0Jp_nU_nS5dfssPQdtc7h';

        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

        const client = Client.initWithMiddleware({
            authProvider: {
                getAccessToken: async () => {
                    const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
                    return tokenResponse.token;
                }
            }
        });

        try {
            const chatId = context.activity.conversation.id;
            let allMessages: any[] = [];
            let nextLink = `/chats/${chatId}/messages?$top=50`;

            do {
                // PS_09, PS_10
                const response = await client.api(nextLink)
                    .version('v1.0')
                    .get();

                allMessages = allMessages.concat(response.value);
                nextLink = response['@odata.nextLink'] || null;
            } while (nextLink);

            return allMessages;
        } catch (error) {
            console.error('Error fetching chat history:', error);
            await context.sendActivity('Error fetching chat history');
            return [];
        }
    }

    // PS_23
    async postChatSummary(context: TurnContext, summary: string): Promise<void> {
        debugger
        await context.sendActivity(summary);
        // const card = await openAdaptiveCard("summary");
        // const showActivity = MessageFactory.attachment(CardFactory.adaptiveCard(card));
        // showActivity.attachments = [CardFactory.adaptiveCard(card)];
        // await context.sendActivity(showActivity);
    }

    // PS_26 - PS_32
    async uploadAndSendPDF(context: TurnContext, summary: string): Promise<void> {
        debugger
        try {
            const pdfBytes = await createPDF(summary);
            const downloadLink = await uploadPDFToSharePoint(pdfBytes);
            await context.sendActivity({
                text: `[DownLoad](${downloadLink})`,
                type: ActivityTypes.Message
            });
        } catch (error) {
            console.error('Error uploading PDF to SharePoint:', error);
            await context.sendActivity('Error uploading PDF to SharePoint');
        }
    }
}
