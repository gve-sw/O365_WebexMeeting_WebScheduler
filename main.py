'''
Copyright (c) 2020 Cisco and/or its affiliates.

This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at

               https://developer.cisco.com/docs/licenses

All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
'''

import yaml, datetime, string, random, icalendar, requests, urllib
from xmltodict import parse as xml_to_dict
from flask import Flask, request, redirect, url_for, render_template

# global variables
o365_access_token = None
webex_access_token = None
webex_username = None
webex_session_ticket = None
redirected = None
o365_groups = {}
o365_owner = {}
meeting_data = {}
config = yaml.safe_load(open("credentials.yml"))

MS_LOGIN_API_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0".format(tenant=config['azure_client_tenant'])
MS_GRAPH_API_URL = "https://graph.microsoft.com"
WEBEX_LOGIN_API_URL = "https://webexapis.com/v1"
WEBEX_MEETINGS_API_URL = "https://api.webex.com/WBXService/XMLService"

# Flask app
app = Flask(__name__)

# to retrieve a Webex Meeings XML API session ticket from a Webex access token
def webex_meetings_session_ticket(webex_username):
    session_ticket_xml = """
    <?xml version="1.0" encoding="UTF-8"?>
        <serv:message xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <header>
                <securityContext>
                    <webExID>{webex_username}</webExID>
                    <siteName>{webex_site}</siteName>
                </securityContext>
            </header>
            <body>
                <bodyContent xsi:type="java:com.webex.service.binding.user.AuthenticateUser">
                    <accessToken>{webex_access_token}</accessToken>
                </bodyContent>
            </body>
        </serv:message>
    """
    data = session_ticket_xml.format(webex_username=webex_username,
                                     webex_site=config['webex_site'],
                                     webex_access_token=webex_access_token)
    get_session_ticket = requests.post(WEBEX_MEETINGS_API_URL, data=data)
    get_session_ticket_text = xml_to_dict(get_session_ticket.text)
    global webex_session_ticket
    webex_session_ticket = get_session_ticket_text['serv:message']['serv:body']['serv:bodyContent']['use:sessionTicket']
    return webex_session_ticket

# to get Webex host permissions from a user
def webex_host_permissions():
    host_permission_xml = '''
        <?xml version="1.0" encoding="UTF-8"?>
        <serv:message xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <header>
                <securityContext>
                    <webExID>{webex_username}</webExID>
                    <sessionTicket>{webex_session_ticket}</sessionTicket>
                    <siteName>{webex_site_name}</siteName>
                </securityContext>
            </header>
            <body>
                <bodyContent xsi:type="java:com.webex.service.binding.user.GetUser">
                        <webExId>{webex_username}</webExId>
                </bodyContent>
            </body>
        </serv:message>
        '''
    body = host_permission_xml.format(webex_username=webex_username,
                                      webex_session_ticket=webex_session_ticket,
                                      webex_site_name=config['webex_site'])
    request = requests.post(WEBEX_MEETINGS_API_URL, data=body)
    d = xml_to_dict(request.text)
    owner_choice_webex = []
    try:
        host_permissions = d['serv:message']['serv:body']['serv:bodyContent']['use:scheduleFor']['use:webExID']
        if isinstance(host_permissions, str): # if only one, the email address is stored in a string
            owner_choice_webex.append(host_permissions)
        elif isinstance(host_permissions, list): # if multiple, email addresses are stored in a list of strings
            for webex_email in host_permissions:
                owner_choice_webex.append(webex_email)
    except: # if none, an error will be thrown
        pass
    return owner_choice_webex


# to create the Webex Meetings XML body, depending on whether the meeting is repeated or not
def create_meetings_xml(input_repeatmeeting_pattern):
    create_meetings_xml_pt1 = """<?xml version="1.0" encoding="UTF-8"?>
        <serv:message xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <header>
                <securityContext>
                    <webExID>{webex_username}</webExID>
                    <sessionTicket>{webex_session_ticket}</sessionTicket>
                    <siteName>{webex_site_name}</siteName>
                </securityContext>
            </header>
            <body>
                <bodyContent
                    xsi:type="java:com.webex.service.binding.meeting.CreateMeeting">
                    <accessControl>
                        <meetingPassword>{meeting_password}</meetingPassword>
                    </accessControl>
                    <metaData>
                        <confName>{meeting_name}</confName>
                        <meetingType>3</meetingType>
                        <agenda>{meeting_agenda}</agenda>
                    </metaData>
                    <participants>
                        <maxUserNumber>100</maxUserNumber>
                    </participants>
                    <enableOptions>
                        <chat>true</chat>
                        <poll>true</poll>
                        <audioVideo>true</audioVideo>
                        <supportE2E>false</supportE2E>
                        <autoRecord>false</autoRecord>
                    </enableOptions>
                    <schedule>
                        <startDate>{start_date}</startDate>
                        <openTime>900</openTime>
                        <joinTeleconfBeforeHost>false</joinTeleconfBeforeHost>
                        <duration>{duration_minutes}</duration>
                        <timeZoneID>22</timeZoneID>
                        <hostWebExID>{owner}</hostWebExID>
                    </schedule>
                    """
    create_meetings_xml_pt3 = """
                </bodyContent>
            </body>
        </serv:message>
    """
    if input_repeatmeeting_pattern == None:
        create_meetings_xml = create_meetings_xml_pt1 + create_meetings_xml_pt3
    elif input_repeatmeeting_pattern == "daily":
        create_meetings_xml_pt2 = """
                    <repeat>
                        <repeatType>{pattern}</repeatType>
                        <interval>1</interval>
                    </repeat>
        """
        create_meetings_xml = create_meetings_xml_pt1 + create_meetings_xml_pt2 + create_meetings_xml_pt3
    elif input_repeatmeeting_pattern == "weekly":
        create_meetings_xml_pt2 = """
                    <repeat>
                        <repeatType>{pattern}</repeatType>
                        <interval>1</interval>
                        <dayInWeek>
                            <day>{dayInWeek}</day>
                        </dayInWeek>
                    </repeat>
        """
        create_meetings_xml = create_meetings_xml_pt1 + create_meetings_xml_pt2 + create_meetings_xml_pt3
    elif input_repeatmeeting_pattern == "monthly":
        create_meetings_xml_pt2 = """
                    <repeat>
                        <repeatType>{pattern}</repeatType>
                        <interval>1</interval>
                        <dayInMonth>{dayInMonth}</dayInMonth>
                    </repeat>
        """
        create_meetings_xml = create_meetings_xml_pt1 + create_meetings_xml_pt2 + create_meetings_xml_pt3
    elif input_repeatmeeting_pattern == "yearly":
        create_meetings_xml_pt2 = """
                    <repeat>
                        <repeatType>{pattern}</repeatType>
                        <monthInYear>{monthInYear}</monthInYear>
                        <dayInMonth>{dayInMonth}</dayInMonth>
                    </repeat>
        """
        create_meetings_xml = create_meetings_xml_pt1 + create_meetings_xml_pt2 + create_meetings_xml_pt3

    return create_meetings_xml


# login page
@app.route('/')
def mainpage_login():
    global redirected
    redirected = None
    return render_template('mainpage_login.html')


# login redirects to Webex for user to provide login data
@app.route('/webexlogin', methods=['POST'])
def webexlogin():
    WEBEX_USER_AUTH_URL = WEBEX_LOGIN_API_URL + "/authorize?client_id={client_id}&response_type=code&redirect_uri={redirect_uri}&response_mode=query&scope={scope}".format(
        client_id=urllib.parse.quote(config['webex_integration_client_id']),
        redirect_uri=urllib.parse.quote(config['webex_integration_redirect_uri']),
        scope=urllib.parse.quote(config['webex_integration_scope'])
    )

    return redirect(WEBEX_USER_AUTH_URL)


# based on Webex login information, a Webex access token is retrieved
@app.route('/webexoauth', methods=['GET'])
def webexoauth():
    webex_code = request.args.get('code')

    headers_token = {
        "Content-type": "application/x-www-form-urlencoded"
    }
    body = {
        'client_id': config['webex_integration_client_id'],
        'code': webex_code,
        'redirect_uri': config['webex_integration_redirect_uri'],
        'grant_type': 'authorization_code',
        'client_secret': config['webex_integration_client_secret']
    }
    get_token = requests.post(WEBEX_LOGIN_API_URL + "/access_token?", headers=headers_token, data=body)

    global webex_access_token
    webex_access_token = get_token.json()['access_token']

    return redirect(url_for('.o365login'))


# login to O365, same workflow as with Webex (will not be visible to user if Webex SSO login uses Microsoft Azure as IdP)
@app.route("/o365login")
def o365login():
    MS_USER_AUTH_URL = MS_LOGIN_API_URL + "/authorize?client_id={client_id}&response_type=code&redirect_uri={redirect_uri}&response_mode=query&scope={scope}".format(
        client_id=config['azure_client_id'],
        redirect_uri=config['azure_client_redirect_uri'],
        scope=config['azure_permissions'])

    return redirect(MS_USER_AUTH_URL)


@app.route('/o365oauth', methods=['GET'])
def o365_oauth():
    o365_code = request.args.get('code')

    headers_token = {
        "Content-type": "application/x-www-form-urlencoded"
    }
    body = {
        'client_id': config['azure_client_id'],
        'scope': config['azure_permissions'],
        'code': o365_code,
        'redirect_uri': config['azure_client_redirect_uri'],
        'grant_type': 'authorization_code',
        'client_secret': config['azure_client_secret']
    }
    get_token = requests.post(MS_LOGIN_API_URL + "/token?", headers=headers_token, data=body)

    global o365_access_token
    o365_access_token = get_token.json()['access_token']

    return redirect(url_for('.mainpage'))


# main page for the user to provide meeting information in the HTML form
@app.route('/mainpage')
def mainpage():
    # to get the username of the Webex user, here equal to email address
    global webex_username
    webex_me_details = requests.get('https://webexapis.com/v1/people/me', headers={'Authorization': 'Bearer ' + webex_access_token}).json()
    webex_username = webex_me_details['emails'][0]

    # to get a session ticket from the Webex access token and that is required for using the Webex XML API
    global webex_session_ticket
    webex_session_ticket = webex_meetings_session_ticket(webex_username)

    # to collect information for the meeting form that are based on the user's O365 and Webex permissions
    headers_group = {
        "Authorization": "Bearer " + o365_access_token
    }

    # to populate the required and optional participant field in the HTML form, based on O365 email groups
    global o365_groups
    group_choice = []
    o365_groups = requests.get(MS_GRAPH_API_URL + "/v1.0/groups", headers=headers_group)
    for group in o365_groups.json()['value']:
        email = group['mail']
        group_choice.append(email)

    # to populate the meeting host/owner field in the HTML form, requirement: user must have editing rights to the O365 calendar and Webex scheduling permissions
    global o365_owner
    o365_owner = requests.get(MS_GRAPH_API_URL + "/v1.0/me/calendars", headers=headers_group).json()
    owner_choice_o365 = []
    for calendar in o365_owner['value']:
        if calendar['canEdit'] == True:
            owner_choice_o365_email = calendar['owner']['address']
            owner_choice_o365.append(owner_choice_o365_email)
    # to only allow owners/hosts of calendars users have both O365 edit rights and Webex scheduling permissions
    owner_choice_webex = webex_host_permissions()
    owner_choice = [webex_username] # own calendar is always an option
    for o365_allowed_email in owner_choice_o365:
        for webex_allowed_email in owner_choice_webex:
            if o365_allowed_email == webex_allowed_email:
                z = webex_allowed_email
                owner_choice.append(z)

    # to check if it is a redirect from a submitted form
    if redirected == "success":
        alert = 1
    elif redirected == "failure":
        alert = 2
    else:
        alert = None

    return render_template('mainpage.html', group_choice=group_choice, owner_choice=owner_choice, alert=alert)


# to retrieve the information from the HTML form after the form is submitted
@app.route('/submit', methods=['POST'])
def submit():
    req = request.form

    # to check whether checkboxes were ticked in the HTML form and store the values accordingly
    if "repeatmeeting" in req.keys():
        input_repeatmeeting_pattern = req["pattern"]
    else:
        input_repeatmeeting_pattern = None

    if "notifyrecipients" in req.keys():
        input_recipients_dropdown = req["recipients"]
    else:
        input_recipients_dropdown = None
    if "notifyCCrecipients" in req.keys():
        input_CCrecipients_dropdown = req["CCrecipients"]
    else:
        input_CCrecipients_dropdown = None

    # to make the information globally available
    global meeting_data
    meeting_data = {
        "input_title": req["title"],
        "input_agenda": req["agenda"],
        "input_date": req["date"],
        "input_time_start": req["starttime"],
        "input_time_end": req["endtime"],
        "input_repeatmeeting_pattern": input_repeatmeeting_pattern,
        "input_owner": req['owner'],
        "input_recipients_dropdown": input_recipients_dropdown,
        "input_CCrecipients_dropdown": input_CCrecipients_dropdown
    }

    return redirect(url_for('.invite'))


# to prepare and send the O365 meeting invite, incl. Webex Meetings details
@app.route('/invite', methods=['GET'])
def invite():
    # to get the information required for the O365 and Webex invite and prepare it for the right format
    input_title = meeting_data['input_title']
    input_agenda = meeting_data['input_agenda']
    input_date = meeting_data['input_date']
    input_date_year = int(input_date[:4])
    input_date_month = int(input_date[5:7])
    input_date_day = int(input_date[8:10])
    weekdays = ("MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY")
    input_date_weekday = weekdays[datetime.date(input_date_year, input_date_month, input_date_day).weekday()]
    input_time_start = meeting_data['input_time_start']
    input_time_start_hour = int(input_time_start[:2])
    input_time_start_minute = int(input_time_start[3:])
    input_time_start_outlook = str(input_date + "T" + input_time_start + ":00")
    input_time_start_webex = str(input_date_month) + "/" + str(input_date_day) + "/" + str(input_date_year) + " " + str(input_time_start) + ":00"
    input_time_end = meeting_data['input_time_end']
    input_time_end_hour = int(input_time_end[:2])
    input_time_end_minute = int(input_time_end[3:])
    input_time_end_outlook = str(input_date + "T" + input_time_end + ":00")
    input_repeatmeeting_pattern = meeting_data["input_repeatmeeting_pattern"]
    input_owner = meeting_data["input_owner"]
    input_recipients_dropdown = meeting_data['input_recipients_dropdown']
    input_CCrecipients_dropdown = meeting_data['input_CCrecipients_dropdown']
    input_meeting_duration = datetime.datetime(input_date_year, input_date_month, input_date_day, input_time_end_hour, input_time_end_minute) - datetime.datetime(input_date_year, input_date_month, input_date_day, input_time_start_hour, input_time_start_minute)
    input_meeting_duration_int = int(input_meeting_duration.seconds / 60)

    # to create a valid Webex Meetings password
    meeting_password_criteria = string.ascii_lowercase + "".join([str(i) for i in range(0, 11)])
    meeting_password = ''.join(random.choice(meeting_password_criteria) for i in range(8))

    # to schedule a Webex Meeting
    webex_meetings_xml = create_meetings_xml(input_repeatmeeting_pattern)
    if input_repeatmeeting_pattern == None:
        body = webex_meetings_xml.format(webex_username=webex_username,
                                         webex_session_ticket=webex_session_ticket,
                                         webex_site_name=config['webex_site'],
                                         meeting_password=meeting_password,
                                         meeting_name=input_title,
                                         meeting_agenda=input_agenda,
                                         start_date=input_time_start_webex,
                                         duration_minutes=input_meeting_duration_int,
                                         owner=input_owner
                                         )
    elif input_repeatmeeting_pattern == "daily":
        body = webex_meetings_xml.format(webex_username=webex_username,
                                         webex_session_ticket=webex_session_ticket,
                                         webex_site_name=config['webex_site'],
                                         meeting_password=meeting_password,
                                         meeting_name=input_title,
                                         meeting_agenda=input_agenda,
                                         start_date=input_time_start_webex,
                                         duration_minutes=input_meeting_duration_int,
                                         owner=input_owner,
                                         pattern=input_repeatmeeting_pattern.upper()
                                         )
    elif input_repeatmeeting_pattern == "weekly":
        body = webex_meetings_xml.format(webex_username=webex_username,
                                         webex_session_ticket=webex_session_ticket,
                                         webex_site_name=config['webex_site'],
                                         meeting_password=meeting_password,
                                         meeting_name=input_title,
                                         meeting_agenda=input_agenda,
                                         start_date=input_time_start_webex,
                                         duration_minutes=input_meeting_duration_int,
                                         owner=input_owner,
                                         pattern=input_repeatmeeting_pattern.upper(),
                                         dayInWeek=input_date_weekday
                                         )
    elif input_repeatmeeting_pattern == "monthly":
        body = webex_meetings_xml.format(webex_username=webex_username,
                                         webex_session_ticket=webex_session_ticket,
                                         webex_site_name=config['webex_site'],
                                         meeting_password=meeting_password,
                                         meeting_name=input_title,
                                         meeting_agenda=input_agenda,
                                         start_date=input_time_start_webex,
                                         duration_minutes=input_meeting_duration_int,
                                         owner=input_owner,
                                         pattern=input_repeatmeeting_pattern.upper(),
                                         dayInMonth=input_date_day
                                         )
    elif input_repeatmeeting_pattern == "yearly":
        body = webex_meetings_xml.format(webex_username=webex_username,
                                         webex_session_ticket=webex_session_ticket,
                                         webex_site_name=config['webex_site'],
                                         meeting_password=meeting_password,
                                         meeting_name=input_title,
                                         meeting_agenda=input_agenda,
                                         start_date=input_time_start_webex,
                                         duration_minutes=input_meeting_duration_int,
                                         owner=input_owner,
                                         pattern=input_repeatmeeting_pattern.upper(),
                                         monthInYear=input_date_month,
                                         dayInMonth=input_date_day
                                         )
    meeting_creation_xml = requests.post(WEBEX_MEETINGS_API_URL, data=body)

    # to prepare the Webex Meeting to send in the O365 invite
    if meeting_creation_xml.status_code == requests.codes.ok:
        d = xml_to_dict(meeting_creation_xml.text)

        meeting_password = d['serv:message']['serv:body']['serv:bodyContent']['meet:meetingPassword']

        # Extract meeting link from iCal
        ical_link = d['serv:message']['serv:body']['serv:bodyContent']['meet:iCalendarURL']['serv:host']

        resp_ical = requests.get(ical_link)

        if resp_ical.status_code == requests.codes.ok:
            ics_cal = icalendar.Calendar.from_ical(resp_ical.text)

            # Extract meeting
            for comp in ics_cal.walk():
                if comp.name == "VEVENT":
                    descr = comp.get('X-ALT-DESC')
                    # Add the meeting password
                    descr = descr.replace("Please obtain your meeting password from your host.",
                                          meeting_password)
                    outlook_content = descr

    # to define the headers for O365 API calls
    headers_event = {
        "Content-type": "application/json",
        "Authorization": "Bearer " + o365_access_token
    }

    # to get the email addresses of people as part of the O365 group if chosen as required and/or optional participants in the HTML form
    attendees = []
    if input_recipients_dropdown != None:
        for item in o365_groups.json()['value']:
            if item['mail'] == input_recipients_dropdown:
                group_id = item['id']
                recipient_group_required = requests.get(MS_GRAPH_API_URL + "/v1.0/groups/" + group_id + "/members",
                                               headers=headers_event)
                for mail in recipient_group_required.json()['value']:
                    mail_address = mail['mail']
                    recipient_name = mail['displayName']
                    attendees.append(
                        {
                            "emailAddress": {
                                "address": mail_address,
                                "name": recipient_name
                            },
                            "type": "required"
                        }
                    )
    if input_CCrecipients_dropdown != None:
        for item in o365_groups.json()['value']:
            if item['mail'] == input_CCrecipients_dropdown:
                group_id = item['id']
                recipient_group_optional = requests.get(MS_GRAPH_API_URL + "/v1.0/groups/" + group_id + "/members",
                                               headers=headers_event)
                for mail in recipient_group_optional.json()['value']:
                    mail_address = mail['mail']
                    recipient_name = mail['displayName']
                    attendees.append(
                        {
                            "emailAddress": {
                                "address": mail_address,
                                "name": recipient_name
                            },
                            "type": "optional"
                        }
                    )

    # to prepare the O365 meeting invite body based on the information provided and gathered above
    o365_invite = {
        "subject": input_title,
        "body": {
            "contentType": "HTML",
            "content": input_agenda + "\n" + outlook_content
        },
        "start": {
            "dateTime": input_time_start_outlook,
            "timeZone": "W. Europe Standard Time"
        },
        "end": {
            "dateTime": input_time_end_outlook,
            "timeZone": "W. Europe Standard Time"
        },
        "location": {
            "displayName": "@webex"
        },
        "attendees": attendees,
        "allowNewTimeProposals": True
    }

    # to add recurrence information to the O365 invite if it is a repeated meeting
    if input_repeatmeeting_pattern != None:
        if input_repeatmeeting_pattern == "daily":
            pattern = {
                "type": "daily",
                "interval": 1,
            }
        elif input_repeatmeeting_pattern == "weekly":
            pattern = {
                "type": "daily",
                "interval": 7
            }
        elif input_repeatmeeting_pattern == "monthly":
            pattern = {
                "type": "absoluteMonthly",
                "interval": 1,
                "dayOfMonth": input_date_day,
            }
        elif input_repeatmeeting_pattern == "yearly":
            pattern = {
                "type": "absoluteYearly",
                "interval": 1,
                "dayOfMonth": input_date_day,
                "month": input_date_month
            }
        range_noEnd = {
            "type": "noEnd",
            "startDate": input_date,
        }
        o365_invite['recurrence'] = {"pattern": pattern, "range": range_noEnd}

    # to get the correct calendar to schedule the meeting for, depending on the information of the meeting host/owner in the HTML form
    for calendar in o365_owner['value']:
        if calendar['canEdit'] == True:
            if calendar['owner']['address'] == input_owner:
                calendar_id = calendar['id']

    # to send the API call to create the O365 meeting with the information provided and gathered before
    outlook_invite = requests.post(MS_GRAPH_API_URL + '/v1.0/me/calendars/' + calendar_id +  '/events', headers=headers_event, json=o365_invite)

    # to provide the correct feedback to the user when having submitted the form
    global redirected
    if outlook_invite.status_code == requests.codes.created:
        redirected = "success"
    else:
        redirected = "failure"
    return redirect(url_for('.mainpage'))


if __name__ == "__main__":
    app.run()