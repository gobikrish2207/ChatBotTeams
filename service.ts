import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';
import axios from 'axios';
import { Client } from '@microsoft/microsoft-graph-client';
import qs from 'qs';

// Function to get access token
export async function getAccessToken(): Promise<string> {
    try {
        const tokenUrl = `https://login.microsoftonline.com/cd5dbd24-330a-4174-8e00-dc550f63976b/oauth2/v2.0/token`;
        const data = qs.stringify({
            'client_id': 'e1cd3b9a-0ae7-467c-847f-07f521343b85',
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials',
            'client_secret': '5G.8Q~O0yZ3PSHoFTwF0Jp_nU_nS5dfssPQdtc7h',
        });
        const headers = { 'Content-Type': 'application/x-www-form-urlencoded' };
        const resp = await axios.post(tokenUrl, data, { headers: headers });
        return resp.data.access_token;
    } catch (e) {
        console.log('Access Token Error', e);
        throw new Error('Failed to obtain access token');
    }
}

// Function to get chat summary
// PS_19 - PS_21
export async function getChatSummary(chatText: string): Promise<string> {
    try {
        const body = {
            "messages": [
                {
                    "role": "system",
                    "content": `You are a language model tasked with describing a chat conversation between a group of people on a specific topic. The chat conversation involves multiple people discussing the incident raised in ServiceNow, your task is to: 1. Summarize the key points discussed in the conversation without duplicating any points and provide a description of the chat including all the important points. The output should be a JSON object with one key: 

                    - "summary": A concise summary of the chat conversation without duplicate and list of bulleted points of the summary. 
        
                    <Task>  
        
                    <Instructions>  
        
                    1. Review the provided chat conversation carefully.  
        
                    2. Identify the main problem or issue being discussed.  
        
                    3. Summarize all the key points discussed in the conversation, including the context, problem description, and any relevant information, in a concise manner without duplicating any points but do not skip any content or message.  
        
                    4. Extract the resolution steps mentioned by the support agent or agreed upon during the conversation.  
        
                    5. The summary should always be bulleted points, the summary should at least contain 5 to 1- points in bulleted format but it should not exclude any messages.  
        
                    6. Ensure that the instructions are clear, concise, and easy to follow.  
        
                    7. Remove any personal names, identifiable information, or redundant messages from the conversation. Double-check before providing the summary that all the chat has been considered for the summary and do not skip any points.  
        
                    8. Provide the output as a JSON object with the following key:  
        
                      - "summary": A string containing the summary of the chat conversation without duplicate points and provide a description of the chat including all the important points in bullets.  
        
                    9. Do not include any additional explanations or text outside the JSON object.  
        
                    10. Follow the output format given as <Output> and do not include anything else.  
        
                    11. Each point should be separated by '\\n' as a new separator, do not include anything else.  
        
                    12. Always the summary should be in points no matter what.  
        
                    13. Do not skip any message; consider all the messages while providing the summary. 
        
                    </Instructions>  
        
                    <Output>  
        
                    { "summary": "- summary of the chat" \\n "- summary of the chat" \\n "- summary of the chat" \\n "- summary of the chat" \\n "- summary of the chat" \\n }  
        
                    </Output>`
                },
                {
                    "role": "user",
                    "content": chatText
                }
            ]
        };

        const apiKey = process.env.OPENAI_API_KEY;
        const url = 'https://api.openai.com/v1/chat/completions';
        const headers = { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' };
        const response = await axios.post(url, body, { headers });
        const completion = response.data.choices[0].message.content;

        return completion;
    } catch (error) {
        console.error('Error generating chat summary:', error);
        throw new Error('Failed to generate chat summary');
    }
}

// Function to create PDF
// PS_26 - PS_27
export async function createPDF(summary: string): Promise<Uint8Array> {
    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([600, 400]);
    const timesRomanFont = await pdfDoc.embedFont(StandardFonts.TimesRoman);
    const { width, height } = page.getSize();
    const fontSize = 12;
    const textWidth = timesRomanFont.widthOfTextAtSize(summary, fontSize);
    const textHeight = timesRomanFont.heightAtSize(fontSize);
    const textX = 50;
    const textY = height - 4 * textHeight;

    page.drawText(summary, {
        x: textX,
        y: textY,
        size: fontSize,
        font: timesRomanFont,
        color: rgb(0, 0, 0)
    });

    const pdfBytes = await pdfDoc.save();
    return pdfBytes;
}

// Function to upload PDF to SharePoint
// PS_28 - PS_31
export async function uploadPDFToSharePoint(pdfBuffer: Uint8Array): Promise<string> {
    const accessToken = await getAccessToken();
    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });

    const fileName = `chat_summary_${Date.now()}.pdf`;
    const folderPath = `/Shared Documents`;

    const response = await client.api(`/sites/root/drive/root:${folderPath}/${fileName}:/content`)
        .put(pdfBuffer);

    const itemId = response.id;
    const item = await client.api(`/sites/root/drive/items/${itemId}`).get();
    const downloadUrl = item["@microsoft.graph.downloadUrl"];

    return downloadUrl;
}
