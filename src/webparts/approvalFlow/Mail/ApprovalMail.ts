import "@pnp/sp/lists";
import { IEmailProperties } from "@pnp/sp/presets/all";
import { SPFI } from "@pnp/sp";
import { getSP } from "../service/PnPConfig";
import "@pnp/sp/sputilities";

export async function ApproveMail(
  requesterMail: string,
  requesterName: string,
  userName: string
) {
  try {
    const emailProps: IEmailProperties = {
      To: [requesterMail],
      CC: [],
      BCC: [],
      Subject: "Approval Notification!",
      Body: `
        <html>
        <head>
           <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        </head>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
           <p>Dear <b>${requesterName}</b>,</p>
           <p>Your request has been approved.</p>
           <br/>
           <div>
            <p>Thanks & Regards </p>
            <p><b>${userName}</b></p>
           </div>
       </body>
       </html>
      `,
    };
    const sp: SPFI = await getSP();
    await sp.utility.sendEmail(emailProps);

    console.log("Email sent successfully");
  } catch (error) {
    console.error("Error sending email:", error);
    throw error; // Rethrow the error if needed for further handling
  }
}
