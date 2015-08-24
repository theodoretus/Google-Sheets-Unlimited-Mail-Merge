# Google-Sheets-Unlimited-Mail-Merge
A mail merge that sends any number of emails (within your daily quota), labels them, and tracks replies. Please note that this script was written with specific functional needs in mind, but the script can be altered to suit others' needs.

This mail merge is written using Google Apps Script (GAS). To utilize this script, you will first need to follow these installation steps:

1. Make your own copy of <a href="https://docs.google.com/spreadsheets/d/12d0jan8MQSV9RuM6Sw5xHmizorZcVXJQLT8Z95BVM4w/edit?usp=sharing">this spreadsheet.</a> 
2. Open up the Script Editor (under "Tools") in your new sheet.
3. Copy + paste each of the .gs files within this project into new files within the Script Editor. (Note: Google implicitly links GAS files within projects, thus there in no need for explicit requires in the .gs files). 
4. Alter lines 144 & 237 in sender.gs, as well as lines 72 & 163 in responses.gs, to reflect the email address that you would like failure messages sent to.
5. Disable threading in your Gmail Inbox (to properly count replies), by following <a href="https://support.google.com/mail/answer/5900?hl=en">these instructions.</a>

Use of the spreadsheet should be fairly intuitive. The essential workflow is: input a list of emails and custom info -> validate emails -> fetch a specific draft and prepare boilerplate template -> schedule sending of personalized emails -> allow Google to send, label, and track emails + responses.

***

Breaking down the steps: 

*Input a list of emails and custom info*: this info will be entered into the "Custom Info" sheet. In my sheet, an email, a name (in this case required), and an optional custom message are fields. 

*Validate emails*: Under "Functions", run "1. Validate Emails" to accomplish this. Fix any highlighted entries, then run again to confirm validation.

*Fetch a specific draft and prepare boilerplate template*: To fetch a boilerplate draft, enter the draft's subject line in the "Template Grabber" sheet, then run "2. Get Draft". This subject line will be used as the subject line when actually sending, so be sure to have settled upon a subject line by the time of drafting your boilerplate template. You can then fill in any of the several optional fields that exist for altering aspects of your template (inserting at attachment/inline image from your Drive, changing your sending name, etc).

IMPORTANT: The boilerplate draft template should be an actual Gmail draft, in your draft folder. It should be devoid of attachments or inline images. Any custom fields that will show in the body of the email (IN THIS CASE "NAME" AND "CUSTOM MESSAGE") should be rendered in the draft as ${Name}, ${Custom Message}, etc. The custom message field is expected to appear before the main body of the email (if a custom message is specified) but after the initial greeting, so HTML break tags are inserted before the custom message when it exists. Thus, the following format is recommended: 

```
Dear ${Name}, ${Custom Message} 

[remainder of email]
```

*Schedule sending of personalized emails*: In the "Mail Sender" tab, you can schedule a time for the emails to go out. Then, run the "3. Start Mail Schedule" function to commit your schedule. The remaining workflow step (*Allow sheet to send, label, and track emails + responses*) will happen automatically (updated on a daily basis in the "Responses" sheet).

***

Error handling is built into the code, but if sending or labeling fails, the "Resume Sending" and "Resume Labeling" functions can be used, respectively. If response tracking fails, you may run the "Check for Responses" function (under "Other Functions"). In these cases, it is wise to run "Delete All Triggers" first, to ensure there are no backed up triggers.

In the event of wanting to stop the script, you can run "Delete All Triggers" to prevent future actions. If you need an emergency halt for sending or labeling as it is actively happening, however, you should manually delete any instances of the word "Scheduled", as it is a requirement for action that the word "Scheduled" be present in the "Mail Sender" sheet next to the email address for which a personalized template is generated, sent, and labeled.

***

Thanks are due to Amit Agarwal, whose contributions to GAS open sourcing are a great inspiration. 

Enjoy!
