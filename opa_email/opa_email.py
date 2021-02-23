import gspread
import numpy as np
import os
import sys
import json
import smtplib
import time
from time import sleep
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from email.message import EmailMessage

# Conditional for when a avc column returns '' then Position should be AVC or VC if checker is the VC
# Input is the name of the checker and the column where in if the person has an AVC in-charge
def position_checker(checker, avc):
    if avc != '':
        return 'Associate'
    elif avc == '':
        if checker == 'Lorenzo Mercado':
            return 'Vice Chairperson'
        elif checker == 'Rainier Magsino':
            return 'Executive Vice Chairperson - Externals'
        else:
            return 'Associate Vice Chairperson'

def oic_disclaimer(checker, position, oic):
    if position == "Associate":
        return  "This officer is under the supervision of " + oic + " (AVC, PNP). Should there be any concerns kindly contact the superviser in-charge."
    else:
        return ''

def opa_sendemail():
    # Establishing Access Credentials
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('cso-secret-key.json', scope)
    gc = gspread.authorize(credentials)

    # Connect to OPA sheet
    opa = gc.open('46th Publicity Tracking System').worksheet('OPA') # change string inside gc.open() to the name of spreadsheet

    # Setting up Email Credentials
    key_pass = open('cso-pnp-pass.json')
    pnp_access = json.load(key_pass)
    email_address = pnp_access['csopnp_email']
    email_password = pnp_access['csopnp_apppass']


    # Access OPA gsheet database
    current_time = datetime.now()
    timestamp = current_time.strftime("%Y-%m-%d %H:%M:%S")

    db = np.array(opa.get_all_values())
    db_len = len(db)
    # Counter for the number of emails sent per iteration
    emails_sent = 0
    for i in range(1, db_len):
        rows = [rows for rows in db]
        row = rows[i]

        # Set a variable for each column
        submission_number = row[0]
        submission_timestamp = row[1]
        submission_organization = row[2]
        submission_name = row[3]
        submission_position = row[4]
        submission_email = row[5]
        submission_postingdate = row[7]
        submission_activitynum = row[8]
        submission_activitytitle = row[9]
        submission_status = row[14]
        submission_remarks = row[15]
        submission_checker = row[16]
        submission_avc = row[17]
        submission_checkdate = row[18]
        submission_emailstatus = row[19]
        submission_pnpposition = position_checker(submission_checker, submission_avc)
        submission_disclaimer = oic_disclaimer(submission_checker, submission_pnpposition, submission_avc)


        if submission_timestamp!='' and (submission_status=='Approved' or submission_status=='Pended') and submission_checker!='' and submission_checkdate!= '' and submission_emailstatus[:13]!="Email Sent at":
            
            # Set up Email Message
            msg = EmailMessage()
            msg['Subject'] = f'[PTS] Online Publicity Approval Status ({submission_number})'
            msg['From'] = email_address
            msg['To'] = submission_email
            
            # Text Version Email Message
            opa_txt = f"""
                {submission_name}
                {submission_position}
                {submission_organization}

                Greetings in St. La Salle!

                This email is to inform you that your submission to the 46th Online Publicity Approval Form (OPA) has been processed.
                Please see below for the status of your submission:

                Activity: {submission_activitytitle} ({submission_activitynum})
                Submission #: {submission_number}
                Status: {submission_status}
                Remarks: 
                {submission_remarks}

                Checked By:
                {submission_checker}
                {submission_pnpposition}
                Publicity and Productions

                {submission_disclaimer}
            """
            msg.set_content(opa_txt)

            # HTML Version Email Message
            opa_html = f"""
                <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <link rel="stylesheet" href="OPA_email.css">

                    <!-- fonts -->
                    <link href="https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,100;0,300;0,400;0,700;0,900;1,100;1,300;1,400;1,700;1,900&family=Poppins:ital,wght@0,400;0,500;0,600;0,700;0,800;0,900;1,400;1,500;1,600;1,700;1,800;1,900&display=swap" rel="stylesheet">

                    <title>Online Publicity divacking System</title>
                    <meta name="viewport" content="width=device-width; initial-scale=1.0;">

                <style>

                ::-moz-selection {{color: #fff;background: #044002;}}
                ::selection {{color: #fff;background: #044002;}}

                body {{
                    font-family:'Lato',sans-serif;
                    width:800px;
                    max-width:800px;
                    height:100%;
                    margin:auto;
                    padding:20px;
                }}

                .header-title {{
                    padding:20px;
                    margin:-50px;
                    width:100%;
                    max-width:100%;
                    height:25%;
                    background:#fff;
                    display:inline-block;
                    margin:auto;
                    box-shadow: 0px 5px 4px -4px rgba(0,0,0,0.2); 
                }}
                
                .cso-logo {{
                    width:10%;
                    height:auto;
                    float:left;
                }}

                .header-text {{
                    float:left;
                    margin-left:15px;
                    }}

                .cso {{
                    font-size:26px;
                    font-weight:800;
                    color:#044002;
                    text-transform:uppercase;
                    padding-top:7px;
                    font-family:'Poppins',sans-serif;
                    }}

                .tagline {{
                    color:#599537;
                    font-size:12px;
                    font-weight:600;
                    font-style:italic;
                    font-family:'Lato',sans-serif;
                    }}

                .wrap {{
                    width:800px;
                    max-width:800px;
                    height:75%;
                    max-height:75%;
                    padding:20px;
                }}

                .email-title {{
                    font-size:32px;
                    font-weight:700;
                    padding:20px 0px 20px 40px; /* top right bottom left*/
                    color:#044002;
                    border-left:30px solid #044002;
                    font-family:'Poppins',sans-serif;}}

                .email-body {{
                    width:70%;
                    font-size:15px;
                    line-height:170%;
                    margin:20px auto;
                    }}

                .footer {{
                    width:100%;
                    max-width:800px;
                    height:15%;
                    color:#F5F5F5;
                    background-color:#1B1B1B;
                    display:inline-block;
                    font-size:9px;
                    line-height:150%;
                    padding:20px;}}

                .footer a {{
                    color:#F5F5F5;
                }}
                    
                .footer li {{list-style:none;}}
                .about {{width:50%;float:left;}}

                .fortysixth-logo {{
                    width:70px;
                    height:70px;
                    float:left;
                    margin:10px 0px -10px 10px;
                }}

                .learn-more {{
                    width:25%;
                    float:left;
                }}

                .contact {{
                    width:25%;
                    float:left;
                }}

                .t {{font-size:10px;font-weight:700;}}

                .learn-more img, .contact img {{height:10px;width:10px;margin-right:7px;}}

                .tab{{ margin: 40px;}}
                </style>
                    
                </head>
                <body>
                <!-- Header -->
                <div class="header">
                    <div class="header-title">
                        <img class="cso-logo" src="https://imgur.com/RZ8W7Yy.png">
                    <div class="header-text">
                        <div class="cso">Council of Student Organizations</div>
                        <div class="tagline">48 ORGANIZATIONS | 9 EXECUTIVE TEAMS | 1 CSO</div>
                    </div><!-- header-text -->
                    </div><!-- header-title -->
                    </div><!-- header -->
                    <!-- Body -->
                    <div class="wrap">
                        <div class="email-title">Online Publicity Approval Status</div>

                        <div class="email-body">
                        <span class = "\\\body"> 
                            {submission_name}<br>
                            {submission_position}<br>
                            {submission_organization}<br>
                            <br><br>
                            Greetings in St. La Salle!
                            <br><br>
                            This email is to inform you that your submission to the 46th Online Publicity Approval Form (OPA) has been processed. Please see below for the status of your submission:
                            <br><br><br>
                            <strong>Activity:</strong> {submission_activitytitle} ({submission_activitynum}) <br>
                            <strong>Submission #:</strong> {submission_number} <br>
                            <strong>Status</strong>: <Strong><em>{submission_status}</em></Strong> <br>
                            <strong>Remarks</strong>: <br>
                                <span class="tab">{submission_remarks}</span>
                            <br><br>
                            <strong>Checked By:</strong>
                            <br><br>
                                <span class="tab"><strong>{submission_checker}</strong></span><br>
                                <span class="tab"><em>{submission_pnpposition}</em></span><br>
                                <span class="tab"><em>Publicity and Productions</em></span><br>
                                <br>
                                <span class="tab"><em>{submission_disclaimer}</em></span>
                        </span>
                            </div>
                    </div>
                    <!-- Contact Details -->
                    <div class="footer">
                    <div>
                        <div class="about">
                        <img class="fortysixth-logo" src="https://imgur.com/3SnhxOg.png">
                        <ul style="float:left;">
                            <li class="t">46th Council of Student Organizations</li>
                            <li>Email Template by CSO Publicity and Productions</li>
                            <li>Learn more about us at DLSU-CSO.ORG</li>
                        </ul>
                        </div>


                        <div class="learn-more">
                        <ul>
                            <li class="t">Learn More About CSO</li>
                            <li><img src="https://imgur.com/ipkWJY4.png"> cso@dlsu.edu.ph</li>
                            <li><img src="https://imgur.com/3bnqkMl.png">
                            dlsu-cso.org</li>
                        </ul>
                        </div>


                        <div class="contact">
                        <ul>
                            <li class="t">Connect With Us</li>
                            <li> <img src="https://imgur.com/NrIgx1d.png">fb.com/cso.dlsu</li>
                            <li> <img src="https://imgur.com/GolodRs.png">@dlsucso</li>
                            <li><img src="https://imgur.com/bjJ6gX4.png">linkedin.com/in/cso.dlsu</li>
                            <li></li>
                        </div>

                    </div>
                    </div>
                </body>
                </html>
            """
            msg.add_alternative(opa_html, subtype='html')

            # Send Email
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(email_address, email_password)
                smtp.send_message(msg)

            # update email status
            email_status = "Email Sent at {}".format(timestamp)
            sheet_row  = i+1
            opa.update_cell(sheet_row, 20, email_status)
            emails_sent += 1
            # print report log to console & log file
            with open('opa_email_run_logs.txt', 'a') as logs:
                logs.write(f"[{timestamp}] An email has been sent for submission {submission_number}.\n")
                logs.close()
            print(f"[{timestamp}] An email has been sent for submission {submission_number}.")

    approved = [i for i in np.array(opa.col_values(20)) if i != '']
    total_submissions = [i for i in np.array(opa.col_values(2)) if i != '']  # An array of email_status column values
    total_pending = (len(total_submissions) - len(approved)) - 1 # Total number of pending submissions awaiting to be checked

    # print report log to console & log file
    with open('opa_email_run_logs.txt', 'a') as logs:
        logs.write(f"[{timestamp}] {emails_sent} emails have been sent and {total_pending} submissions are still awaiting approval status.\n")
        logs.close()
    print(f"[{timestamp}] {emails_sent} emails have been sent and {total_pending} submissions are still awaiting approval status.")

# Run Script
try:
    while True:
        opa_sendemail()
        time.sleep(120)

except KeyboardInterrupt:
    print("The OPA emailing system has been interrupted via KeyboardInterrupt.")

#except:
    #print("The OPA emailing system has been interrupted by ", sys.exc_info()[0],".")
