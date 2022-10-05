import pandas as pd
import win32com.client as client
from pathlib import Path

outlook = client.Dispatch('Outlook.Application')

pop = "batch 2.csv"

receivers = pd.read_csv(pop)

homePath = Path.home() / "OneDrive - advance.online\Documents"

try:
    for i, row in receivers.iterrows():
        
        try:
            result = row["result"]
        except KeyError:
            result = None
        
        if result != "Sent":
            email = outlook.CreateItem(0)
            email.To = row["email"]
            email.Subject = "Your Feedback is Invaluable!"
            email.Bcc = "advance.online+979a843e96@invite.trustpilot.com"
            email.SentOnBehalfOfName = 'hello@advance.online'

            email.HTMLBody = rf"""
                Good Afternoon, 
                <br>
                <br>
                According to our records, we have either previously paid or are currently paying you for work completed through our CIS solution, and we would love to hear your feedback! 
                <br>
                <br>
                To enable contractors, such as yourself, to make an informed decision on which Contractor Services provider to choose, we have recently set up our own Trust Pilot page, and we would like to invite you to leave ADVANCE a review! 
                <br>
                <br>
                You will shortly receive a communication from our Trust Pilot email address, with a link to follow to leave a review. If you do have a few moments to spare, we would really appreciate your feedback!
                <br>
                <br>
                If you have already left a review previously, thank you very much and we hope the service being provided to you continues to be world-class! 
                <br>
                <br>
                <b>Kindest regards</b>,
                <br>
                <br>
                <font color='#7098F0'>
                <b>ADVANCE</b>
                <br>
                <br>
                Office: 01244 564 564
                <br>
                Email: hello@advance.online
                <br>
                Visit: <a href='https://www.advance.online/'>www.advance.online</a>
                <br>
                <br>
                <em>Service is important to us, and we value your feedback. Please tell us how we did today by clicking <a href='https://www.google.com/search?rlz=1C1GCEA_enGB894GB894&ei=jRPJX56yDr2p1fAP0piSqAw&q=advance+contracting&gs_ssp=eJzj4tFP1zfMSDYsNsyxLDdgtFI1qDCxME9MNjY1MDG3SEw1tDC3MqhINjSyTE00SE1MszAzSExN9RJOTClLzEtOVUjOzyspSkwuycxLBwARHBbA&oq=advance+contract&gs_lcp=CgZwc3ktYWIQAxgAMgsILhDHARCvARCTAjIICC4QxwEQrwEyCAguEMcBEK8BMgIIADIICC4QxwEQrwEyBQgAEMkDMggILhDHARCvATICCAAyAggAMgIIADoKCAAQsQMQgwEQQzoOCC4QsQMQgwEQxwEQowI6CAgAELEDEIMBOgQIABBDOgUIABCxAzoLCC4QsQMQxwEQowI6CAguEMcBEKMCOgoILhDHARCjAhBDOg0ILhDHARCjAhBDEJMCOgsILhDHARCvARCRAjoQCC4QsQMQxwEQowIQQxCTAjoFCAAQkQI6BwgAELEDEEM6CwguELEDEMcBEK8BOg0ILhCxAxDHARCjAhBDOgoILhDHARCvARAKOgQIABAKOgIILlCFFFjCOGDAQGgCcAF4AIABxQGIAZIYkgEEMS4xN5gBAKABAaoBB2d3cy13aXrAAQE&sclient=psy-ab#lrd=0x487ac350478ae187:0xc129ea0eaf860aee,3,,,https://www.google.com/search?rlz=1C1GCEA_enGB894GB894&ei=jRPJX56yDr2p1fAP0piSqAw&q=advance+contracting&gs_ssp=eJzj4tFP1zfMSDYsNsyxLDdgtFI1qDCxME9MNjY1MDG3SEw1tDC3MqhINjSyTE00SE1MszAzSExN9RJOTClLzEtOVUjOzyspSkwuycxLBwARHBbA&oq=advance+contract&gs_lcp=CgZwc3ktYWIQAxgAMgsILhDHARCvARCTAjIICC4QxwEQrwEyCAguEMcBEK8BMgIIADIICC4QxwEQrwEyBQgAEMkDMggILhDHARCvATICCAAyAggAMgIIADoKCAAQsQMQgwEQQzoOCC4QsQMQgwEQxwEQowI6CAgAELEDEIMBOgQIABBDOgUIABCxAzoLCC4QsQMQxwEQowI6CAguEMcBEKMCOgoILhDHARCjAhBDOg0ILhDHARCjAhBDEJMCOgsILhDHARCvARCRAjoQCC4QsQMQxwEQowIQQxCTAjoFCAAQkQI6BwgAELEDEEM6CwguELEDEMcBEK8BOg0ILhCxAxDHARCjAhBDOgoILhDHARCvARAKOgQIABAKOgIILlCFFFjCOGDAQGgCcAF4AIABxQGIAZIYkgEEMS4xN5gBAKABAaoBB2d3cy13aXrAAQE&sclient=psy-ab'>here</a> to leave us a review on Google</em>
                </font>
                <br>
                <br>
                <img src="{str(homePath / 'signature.png')}">
                <br>
                <br>
                <font style='color:#9C9C9C' face = 'Ariel' size='0.5'>
                The information in this email is confidential, privileged and protected by copyright. It is intended solely for the addressee. If you are not the intended recipient any disclosure, copying or distribution of this email is prohibited and may not be lawful. If you have received this transmission in error, please notify the sender by replying by email and deleting the email from all of your computer systems. The transmission of this email cannot be guaranteed to be secure or unaffected by a virus. The contents may be corrupted, delayed, intercepted or lost in transmission, and Advance cannot accept any liability for errors, omissions or consequences which may arise.
                <br>
                Registered Office; Ground Floor, VISTA, St David's Park, Ewloe, CH5 3DT. Registered in England and Wales.
                </font>
            """
            
            email.Display(True)
            receivers.at[i, "result"] = "Sent"
except KeyboardInterrupt:
    pass

receivers.to_csv(pop, index=False)