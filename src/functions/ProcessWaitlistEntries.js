const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
require('dotenv').config();


// Azure Function that connects to Google Sheets and sends emails via Microsoft Graph API

// This function will be triggered when new rows are added to the Google Sheet
// You can configure it as an HTTP trigger or timer trigger
module.exports = async function (context, req) {
    try {
        // Process right away for real-time email sending
        const newWaitlistUsers = await getNewWaitlistEntries();
        
        if (newWaitlistUsers.length > 0) {
            const graphClient = getGraphClient();
            
            for (const user of newWaitlistUsers) {
                await sendSpikeWelcomeEmail(graphClient, user.email);
                await markUserAsEmailed(user.rowIndex);
                context.log(`Sent welcome email to: ${user.email}`);
            }
            
            context.res = {
                status: 200,
                body: `Successfully sent emails to ${newWaitlistUsers.length} new waitlist entries`
            };
        } else {
            context.res = {
                status: 200,
                body: 'No new waitlist entries to process'
            };
        }
    } catch (error) {
        context.log.error(`Error processing waitlist: ${error.message}`);
        context.res = {
            status: 500,
            body: `Error processing waitlist: ${error.message}`
        };
    }
};

async function getNewWaitlistEntries() {
    // Google Sheets API credentials
    const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0]; // Adjust if your sheet is not the first one
    
    const rows = await sheet.getRows();
    const newUsers = [];
    
    // Assuming column A has emails and column B tracks if email was sent
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const email = row.get('Email'); // Use your actual column header
        const emailSent = row.get('EmailSent'); // Use your actual column header
        
        if (email && (!emailSent || emailSent.toLowerCase() !== 'true')) {
            newUsers.push({
                email: email,
                rowIndex: i
            });
        }
    }
    
    return newUsers;
}

function getGraphClient() {
    // Azure AD app registration credentials
    const credential = new ClientSecretCredential(
        process.env.TENANT_ID,
        process.env.CLIENT_ID,
        process.env.CLIENT_SECRET
    );
    
    return Client.initWithMiddleware({
        authProvider: {
            getAccessToken: async () => {
                const response = await credential.getToken("https://graph.microsoft.com/.default");
                return response.token;
            }
        }
    });
}

async function sendSpikeWelcomeEmail(graphClient, recipientEmail) {
    const htmlContent = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Welcome to the Spike Waitlist</title>
    <link href="https://fonts.googleapis.com/css2?family=Kanit:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Kanit', Arial, sans-serif;
            background-color: #f7f7f7;
            margin: 0;
            padding: 20px;
            color: #404040;
            line-height: 1.5;
        }
        .container {
            max-width: 500px;
            margin: 0 auto;
            background-color: #ffffff;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        }
        .header {
            background-color: #FCB434;
            padding: 25px 20px;
            text-align: center;
        }
        .header h1 {
            color: #ffffff;
            font-size: 24px;
            font-weight: 700;
            margin: 0;
        }
        .content {
            padding: 25px;
            color: #505050;
            font-size: 16px;
        }
        .footer {
            text-align: center;
            padding: 15px;
            background-color: #f0f0f0;
            color: #777;
            font-size: 14px;
        }
        .footer a {
            color: #FCB434;
            text-decoration: none;
            font-weight: 500;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Welcome to Spike!</h1>
        </div>
        
        <div class="content">
            <p>Thanks for joining our waitlist! We can't wait to have you on board. Keep an eye out – something big is coming later this month...</p>
            <span style="display:none; max-height:0px; overflow:hidden;">The future of Track & Field is here.</span>
        </div>
        
        <div class="footer">
            <p>Best regards,<br>The Spike Team</p>
            <p>Need help? <a href="mailto:support@spikeapp.com">Contact support</a></p>
            <p>© 2025 Spike Play Inc. All rights reserved.</p>
        </div>
    </div>
</body>
</html>`;

    const message = {
        subject: "Welcome to Spike!",
        body: {
            contentType: "HTML",
            content: htmlContent
        },
        toRecipients: [
            {
                emailAddress: {
                    address: recipientEmail
                }
            }
        ],
        from: {
            emailAddress: {
                address: "info@spikeplay.com",
                name: "Spike"
            }
        }
    };
    
    // Send the email using the configured sender in Azure
    await graphClient
        .api('/users/info@spikeplay.com/sendMail')
        .post({
            message: message,
            saveToSentItems: true
        });
    
    return true;
}

async function markUserAsEmailed(rowIndex) {
    // Similar Google Sheets connection as in getNewWaitlistEntries
    const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];
    
    const rows = await sheet.getRows();
    rows[rowIndex].set('EmailSent', 'true');
    await rows[rowIndex].save();
}