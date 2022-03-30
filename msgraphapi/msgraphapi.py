#!/usr/bin/python
from __future__ import division
import requests
import json
import datetime
import urllib
import csv
from io import StringIO


class msgraphapi:
    """ graph api class"""

    def __init__(self, credpath):
        """ credpath is the path to a json file containing tenant credentials """
        creds = json.load(open(credpath))
        self.clientid = creds["clientid"]
        self.client_secret = creds["clientsecret"]
        self.login_url = creds["loginurl"]
        self.tenant = creds["tenant"]

        """ Get an access token for the graph.microsoft.com API """
        self.bodyvals = dict(
            client_id=self.clientid,
            client_secret=self.client_secret,
            grant_type='client_credentials',
            resource='https://graph.microsoft.com',
        )

        self.header = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/json'
        }

        self.request_url = self.login_url + self.tenant + '/oauth2/token'
        self.token_response = requests.post(
            self.request_url, data=self.bodyvals, headers=self.header)

        self.access_token2 = self.token_response.json().get('access_token')
        self.token_type = self.token_response.json().get('token_type')
        self.header_params_GMC = {
            'Authorization': self.token_type + ' ' + self.access_token2}

        self.base_url = "https://graph.microsoft.com/v1.0"
        self.gmc_session = requests.Session()
        self.gmc_session.headers = self.header_params_GMC

    def getroleid(self, rolename):
        """ Gets the guid of a role name """

        request_string = f"{self.base_url}/directoryRoles"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        data = data['value']
        objectid = [x['id'] for x in data if x['displayName'] == rolename]

        if len(objectid) > 0:
            return str(objectid[0])
        else:
            return "Role not found"

    def getuseridbyimmutableid(self, immutableid):
        """ Gets the Object ID for a user given the immutable ID """

        url_encoded = urllib.parse.quote_plus(
            f"onPremisesImmutableId eq '{immutableid}'")
        request_string = f"{self.base_url}/users/?$filter={url_encoded}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return data['value'][0]['id']

    def getuserupnbyimmutableid(self, immutableid):
        """ Gets the UPN for  auser given the immutable ID """

        url_encoded = urllib.parse.quote_plus(
            f"onPremisesImmutableId eq '{immutableid}'")
        request_string = f"{self.base_url}/users/?$filter={url_encoded}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return data['value'][0]['userPrincipalName']

    def getupnid(self, upn):
        """ Gets the guid of a user given the UPN name """

        url_encoded = urllib.parse.quote_plus(f"{upn}")
        request_string = f"{self.base_url}/users/{url_encoded}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        if 'id' in data.keys():
            return data['id']
        else:
            return "Notfound"

    def getupnfromguid(self, guid):
        """ Gets the upn of a user given the guid """

        request_string = f"{self.base_url}/users/{guid}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return data['userPrincipalName']

    def getuserattr(self, userguid):
        """ returns deleted user data based on guid """

        userguid = userguid.replace('-', '')
        request_string = f"{self.base_url}/directory/deletedItems/microsoft.graph.user?&$filter=id eq '{userguid}'"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return data['value'][0]

    def getuserdetails(self, userguid):
        """ returns detailed user data based on guid """

        request_string = f"{self.base_url}/users/{userguid}?$select=businessPhones,department,DisplayName,GivenName,mobilePhone,Surname,JobTitle,mail,manager"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return data

    def updateuserdetails(
            self,
            userguid,
            businessPhones,
            department,
            displayName,
            givenName,
            mobilePhone,
            surname,
            jobTitle,
            mail):
        """ Updates the metadata for a user given the guid """

        request_string = f"{self.base_url}/users/{userguid}"
        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_body = json.dumps({
            "businessPhones": businessPhones,
            "department": department,
            "displayName": displayName,
            "givenName": givenName,
            "mobilePhone": mobilePhone,
            "surname": surname,
            "jobTitle": jobTitle,
            "mail": mail
        })
        response = requests.patch(
            request_string,
            data=request_body,
            headers=header)
        return response

    def getuserexists(self, upn):
        """ Check if a UPN exists in a tenant """

        url_encoded = urllib.parse.quote_plus(f"userPrincipalName eq '{upn}'")
        request_string = f"{self.base_url}/users?&$filter={url_encoded}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        datal = len(data['value'])
        if datal == 0:
            return False
        elif datal == 1:
            return True
        else:
            return "Error"

    def listroles(self):
        """ List Office 365 roles """

        request_string = f"{self.base_url}/directoryRoles"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        roles = [x['displayName'] for x in data['value']]
        return roles

    def addusertorole(self, userguid, roleguid):
        """ Adds a user to a role given the userguid and role guid """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/directoryRoles/{roleguid}/members/$ref"
        request_body = json.dumps({
            "@odata.id": f"{self.base_url}/directoryObjects/{userguid}"
        })
        response = requests.post(
            request_string,
            data=request_body,
            headers=header)
        if response.ok:
            return "Sucess"
        else:
            return f"Failed to add userguid: {userguid}"

    def removeuserfromrole(self, userguid, roleguid):
        """ Adds a user to a role given the userguid and role guid """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/directoryRoles/{roleguid}/members/{userguid}/$ref"
        response = requests.delete(request_string, headers=header)
        if response.ok:
            return "Sucess"
        else:
            return f"Failted to remove userguid: {userguid}"

    def listrolemembers(self, roleguid):
        """ lists the members assigned to a role guid """

        request_string = f"{self.base_url}/directoryRoles/{roleguid}/members"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        data = data['value']
        roles = [x['userPrincipalName'] for x in data]
        return roles

    def auditlogbyid(self, directoryid):
        """ Provides (or gets) a specific Azure Active Directory audit log item """

        request_string = f"{self.base_url}/auditLogs/directoryAudits{directoryid}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return json.dumps(data)

    def auditlog(self, numofdays):
        """ Provides the list of audit logs generated by Azure Active Directory """

        startdate = datetime.date.strftime(
            datetime.date.today() -
            datetime.timedelta(
                days=numofdays),
            '%Y-%m-%d')
        request_string = f"{self.base_url}/auditLogs/directoryAudits?&$filter=activityDateTime ge {startdate}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()

        return json.dumps(data, indent=4, sort_keys=True)

    def getroletemplates(self):
        """ get role template """
        request_string = f"{self.base_url}/directoryRoles"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return json.dumps(data, indent=4, sort_keys=True)

    def activaterole(self, roleid):
        """ Activates directory roles """
        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/directoryRole"
        request_body = json.dumps({
            "roleTemplateId": roleid
        })
        response = requests.post(
            request_string,
            data=request_body,
            headers=header)
        data = response.json()
        return json.dumps(data, indent=4, sort_keys=True)

    def addmembertorole(self, upn, roleguid):
        """ Adds a member to a role """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/directoryRoles/{roleguid}/members/$ref"
        userguid = self.getupnid(upn)
        request_body = json.dumps({
            "id": userguid
        })
        response = requests.post(
            request_string,
            data=request_body,
            headers=header)
        data = response.json()
        return json.dumps(data, indent=4, sort_keys=True)

    def updatememberuser(self, guid):
        """ This changes a guest user to a member of the directory """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/users/{guid}"
        request_body = json.dumps({
            "userType": "Member"
        })
        response = requests.patch(
            request_string,
            data=request_body,
            headers=header)
        return response

    def skuinuse(self, skuid):
        """ For licensing reporting, gets the number of units used for a skuid """

        request_string = f"{self.base_url}/subscribedSkus"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        for skus in data['value']:
            if skus['id'] == skuid:
                total = skus['prepaidUnits']['enabled']
                inuse = skus['consumedUnits']
                percentinuse = (inuse / total * 100)
        return inuse, total, percentinuse

    def getsignins(self):
        """ Gets signins from the audit logs """

        request_string = f"{self.base_url}/auditLogs/signIns"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        users = []
        for user in data['value']:
            users.append(user['userPrincipalName'])
        return users

    def getdirobject(self, directoryid):
        request_string = f"{self.base_url}/directoryObjects/{directoryid}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return data

    def getgroupnamefromid(self, directoryid):
        obj = self.getdirobject(directoryid)
        groupname = obj["displayName"]

        return groupname

    def listmembers(self, directoryid):
        """ Gets the UPN's of users of a group """

        request_string = f"{self.base_url}/groups/{directoryid}/members"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()

        member_list = data['value']
        next_url = ''
        while True:
            if '@odata.nextLink' in data:
                if data['@odata.nextLink'] == next_url:
                    break
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                member_list += next_data['value']
                data = next_data
            else:
                break
        membersupn = [x['userPrincipalName']
                      for x in member_list if 'userPrincipalName' in x]
        return membersupn

    def listmembersid(self, directoryid):
        """ Gets the id's of users of a group """

        request_string = f"{self.base_url}/groups/{directoryid}/members"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()

        member_list = data['value']
        next_url = ''
        while True:
            if '@odata.nextLink' in data:
                if data['@odata.nextLink'] == next_url:
                    break
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                member_list += next_data['value']
                data = next_data
            else:
                break
        membersid = [x['id'] for x in member_list if 'userPrincipalName' in x]
        return membersid

    def listtransitivemembersupn(self, directoryid):
        """ Gets the transitive UPN's of users of a group """

        request_string = f"{self.base_url}/groups/{directoryid}/transitiveMembers/?$select=userPrincipalName"

        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        print(data)
        member_list = data['value']
        next_url = ''
        while True:
            if '@odata.nextLink' in data:
                if data['@odata.nextLink'] == next_url:
                    break
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                member_list += next_data['value']
                data = next_data
            else:
                break
        membersid = [x['userPrincipalName'] for x in member_list]
        return membersid

    def listtransitivemembersid(self, directoryid):
        """ Gets the transitive id's of users of a group """

        request_string = f"{self.base_url}/groups/{directoryid}/transitiveMembers/?$select=id"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()

        member_list = data['value']
        next_url = ''
        while True:
            if '@odata.nextLink' in data:
                if data['@odata.nextLink'] == next_url:
                    break
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                member_list += next_data['value']
                data = next_data
            else:
                break
        membersid = [x['id'] for x in member_list]
        return membersid

    def getgroupdeletedate(self):
        """ Gets all Office 365 groups deleted in the last 30 days """

        request_string = f"{self.base_url}/directory/deletedItems/microsoft.graph.group"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        groupdict = {}
        for group in data['value']:
            deletedate = group['deletedDateTime']
            deletedate = deletedate.split("T")[0]
            dt = datetime.datetime.strptime(deletedate, '%Y-%m-%d')
            pdt = dt + datetime.timedelta(days=30)
            dts = str(pdt)
            dts = dts.split(" ")[0]
            pdtstr = str(dts)
            groupdict.update({group['displayName']: pdtstr})
            sorted_g = sorted(groupdict.items(), key=lambda x: x[1])
        return json.dumps(sorted_g, indent=4)

    def geto365groups(self):
        """ Gets the GUIDs of all Office 365 groups """

        request_string = f"{self.base_url}/groups?$filter=groupTypes/any(c:c+eq+'Unified')"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        groups = []
        for group in data['value']:
            groups.append(group['id'])
        return groups

    def geto365groupowner(self, directoryid):
        """ Gets all office 365 groups without owners """

        request_string = "https://graph.microsoft.com/v1.0/groups/" + directoryid + "/owners"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        owners = []
        try:
            for owner in data['value']:
                owners.append(owner['userPrincipalName'])
        except Exception as error:
            print(error)
            pass
        return owners

    def inviteuser(self, usermail, url):
        """ Invites a guest to a tenant without sending an email """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/invitations"
        request_body = json.dumps({
            "invitedUserEmailAddress": usermail,
            "inviteRedirectUrl": url,
            "sendInvitationMessage": False
        })
        response = requests.post(
            request_string,
            data=request_body,
            headers=header)
        data = response.json()
        return json.dumps(data, indent=4, sort_keys=True)

    def addmembertogroup(self, upn, groupguid):
        """ Adds a user to a group based on UPN """

        userguid = self.getupnid(upn)
        self.addguidtogroup(userguid, groupguid)

    def addguidtogroup(self, userguid, groupguid):
        """ Adds a user guid to a group guid """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/groups/{groupguid}/members/$ref"
        request_body = json.dumps({
            "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{userguid}"
        })
        response = requests.post(
            request_string,
            data=request_body,
            headers=header)
        return response

    def removememberfromgroup(self, upn, groupguid):
        """ Removes a user UPN from a group """

        userguid = self.getupnid(upn)

        self.removeguidfromgroup(userguid, groupguid)

    def removeguidfromgroup(self, userguid, groupuid):
        """ Remove a user guid from a group """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/groups/{groupuid}/members/{userguid}/$ref"
        response = requests.delete(request_string, headers=header)
        return response

    def getguestusers(self):
        """ Gets all guest users in a tenant """

        request_string = f"{self.base_url}/users?$filter=userType eq 'Guest'"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()

        member_list = data['value']
        next_url = ''
        while True:
            if '@odata.nextLink' in data:
                if data['@odata.nextLink'] == next_url:
                    break
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                member_list += next_data['value']
                data = next_data
            else:
                break
        membersupn = [x['userPrincipalName']
                      for x in member_list if 'userPrincipalName' in x]
        return membersupn

    def getguestusersbyid(self):
        """ Gets all guest users in a tenant """

        request_string = f"{self.base_url}/users?$filter=userType eq 'Guest'"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()

        member_list = data['value']
        next_url = ''
        while True:
            if '@odata.nextLink' in data:
                if data['@odata.nextLink'] == next_url:
                    break
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                member_list += next_data['value']
                data = next_data
            else:
                break
        membersid = [x['id'] for x in member_list if 'id' in x]
        return membersid

    def deleteuser(self, userid):
        """ Deletes a user from a tenant """

        url_encoded = urllib.parse.quote_plus(userid)
        request_string = f"{self.base_url}/users/{url_encoded}/"
        response = requests.delete(
            request_string, headers=self.header_params_GMC)
        return response

    def getchanneldata(self, teamid, channelid):
        """ Gets teams channel data """

        request_string = f"{self.base_url}/teams/{teamid}/channels/{channelid}"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return data

    def getchannels(self, teamid):
        """ Gets teams channel data """

        request_string = f"{self.base_url}/teams/{teamid}/channels"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        return data

    def posttochannel(self, teamid, channelid):
        """ Posts to  channel """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/teams/{teamid}/channels/{channelid}/chatThreads"
        request_body = json.dumps({
            "rootMessage": {
                "body": {
                    "contentType": 2,
                    "content": "Hello world"
                }
            }
        })
        response = requests.post(
            request_string,
            data=request_body,
            headers=header)
        data = response.json()
        return json.dumps(data, indent=4, sort_keys=True)

    def getlastsign(self, upn):
        """ Gets signins from the audit logs """

        request_string = f"{self.base_url}/auditLogs/signIns?&$filter=userPrincipalName eq '{upn}'"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        signins = data['value']
        return signins

    def getloggedin(self, upn):
        """ Gets signins from the audit logs """

        request_string = f"{self.base_url}/auditLogs/signIns?&$filter=userPrincipalName eq '{upn}'"
        response = requests.get(request_string, headers=self.header_params_GMC)
        data = response.json()
        try:
            signins = data['value']
            if len(signins) == 0:
                return False
            else:
                return True
        except Exception as e:
            pass
            return f"{upn},error"

    def getgroupid(self, groupname):
        """ Gets the ID of a group from it's displayname """

        request_string = f"{self.base_url}/groups?$filter=displayName eq '{groupname}'"
        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        response = requests.get(request_string, headers=header)
        data = response.json()
        return data['value'][0]['id']

    def getsubs(self):
        """ Gets a download URL for the office activation detail for the tenant """

        request_string = "{self.base_url}/reports/getOffice365ActivationsUserDetail"
        response = requests.get(request_string, headers=self.header_params_GMC)
        return response.url

    def enumerate_csv_report_response(self, response):
        report_list = []
        reader = csv.reader(StringIO(response))
        columns = []
        for index, row in enumerate(reader):
            if index == 0:
                columns = row
            else:
                row_dict = {}
                for indx, column in enumerate(columns):
                    row_dict[column] = row[indx]
                report_list.append(row_dict)
        return report_list

    def getmicrosoft365servicesusercounts(self, period: str = "D7"):
        """ GET /reports/getOffice365ServicesUserCounts(period='D7') """

        request_string = f"{self.base_url}/reports/getOffice365ServicesUserCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getmicrosoft365usercounts(self, period: str = "D7"):
        """ GET /reports/getOffice365ActiveUserCounts(period='D7') """
        
        request_string = f"{self.base_url}/reports/getOffice365ActiveUserCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getmicrosoft365userdetailsbyperiod(self, period: str = "D7"):
        """ GET /reports/getOffice365ActiveUserDetail(period='D7') """

        request_string = f"{self.base_url}/reports/getOffice365ActiveUserDetail(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getmicrosoft365userdetailsbydate(self, date):
        """ GET /reports/getOffice365ActiveUserDetail(date=YYYY-MM-DD) """

        request_string = f"{self.base_url}/reports/getOffice365ActiveUserDetail(date={date})"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    """ EMAIL ACTIVITY REPORTS """

    def getemailactivityusercounts(self, period: str = "D7"):
        """ GET /reports/getEmailActivityUserCounts(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getEmailActivityUserCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getemailactivitycounts(self, period: str = "D7"):
        """ GET /reports/getEmailActivityCounts(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getEmailActivityCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    """ MAILBOX REPORTS """

    def getmailboxusagestorage(self, period: str = "D7"):
        """ GET /reports/getMailboxUsageStorage(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getMailboxUsageStorage(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getmailboxusagequotastatusmailboxcounts(self, period: str = "D7"):
        """ GET /reports/getMailboxUsageQuotaStatusMailboxCounts(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getMailboxUsageQuotaStatusMailboxCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getmailboxusagecounts(self, period: str = "D7"):
        """ GET /reports/getMailboxUsageMailboxCounts(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getMailboxUsageStorage(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getmailboxusagedetail(self, period: str = "D7"):
        """ GET /reports/getMailboxUsageDetail(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getMailboxUsageStorage(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    """ EMAIL APP REPORTS """

    def getemailappuserdetailbyperiod(self, period: str = "D7"):
        """ GET /reports/getEmailAppUsageUserDetail(period='{period}') """

        request_string = f"{self.base_url}/reports/getEmailAppUsageUserDetail(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getemailappuserdetailbydate(self, date):
        """ GET /reports/getEmailAppUsageUserDetail(date='{date}') """

        request_string = f"{self.base_url}/reports/getEmailAppUsageUserDetail(date={date})"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getemailappusageappsusercounts(self, period: str = "D7"):
        """ GET /reports/getEmailAppUsageAppsUserCounts(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getEmailAppUsageAppsUserCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getemailappusageusercounts(self, period: str = "D7"):
        """ GET /reports/getEmailAppUsageUserCounts(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getEmailAppUsageUserCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getemailappusageversionsusercounts(self, period: str = "D7"):
        """ GET /reports/getEmailAppUsageVersionsUserCounts(period='{period_value}') """

        request_string = f"{self.base_url}/reports/getEmailAppUsageVersionsUserCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    """ TEAMS REPORTS """

    def getteamsuseractivitycounts(self, period: str = "D7"):
        """ GET /reports/getTeamsUserActivityCounts(period='D7') """

        request_string = f"{self.base_url}/reports/getTeamsUserActivityCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getteamsusercounts(self, period: str = "D7"):
        """ GET /reports/getTeamsUserActivityUserCounts(period='D7') """

        request_string = f"{self.base_url}/reports/getTeamsUserActivityUserCounts(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getteamsuserdetailbyperiod(self, period: str = "D7"):
        """ GET /reports/getTeamsUserActivityUserDetail(period='D7') """

        request_string = f"{self.base_url}/reports/getTeamsUserActivityUserDetail(period='{period}')"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getteamsuserdetailbydate(self, date):
        """ GET /reports/getTeamsUserActivityUserDetail(date=YYYY-MM-DD) """

        request_string = f"{self.base_url}/reports/getTeamsUserActivityUserDetail(date={date})"

        try:
            response = self.gmc_session.get(request_string).text
        except Exception as ex:
            return ex

        report_list = self.enumerate_csv_report_response(response)

        return report_list

    def getusersdelta(self, delta=None):
        """ Get incremental change for users """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        if delta is None:
            request_string = f"{self.base_url}/users/delta?$select=displayName,givenName,surname,userPrincipalName,mail,manager"
        else:
            request_string = delta
        response = requests.get(request_string, headers=header)
        data = response.json()
        userlist = data['value']
        while True:
            if '@odata.nextLink' in data:
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                userlist += next_data['value']
                data = next_data
            elif '@odata.deltaLink' in data:  # This is the final page of users
                deltatoken = data['@odata.deltaLink']
                break
            else:
                break
        return userlist, deltatoken

    def getgroupdelta(self, delta=None):
        """ Get incremental change for groups """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        if delta is None:
            request_string = f"{self.base_url}/groups/delta?$select=members"
        else:
            request_string = delta
        response = requests.get(request_string, headers=header)
        data = response.json()
        groupdata = data['value']
        while True:
            if '@odata.nextLink' in data:
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                groupdata += next_data['value']
                data = next_data
            elif '@odata.deltaLink' in data:  # This is the final page of users
                deltatoken = data['@odata.deltaLink']
                break
            else:
                break
        return groupdata, deltatoken

    def getgroupdeltanostate(self):
        """ Get a delta token with no state """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/groups/delta?$select=members&$deltaToken=latest"
        response = requests.get(request_string, headers=header)
        data = response.json()
        deltatoken = data['@odata.deltaLink']
        return deltatoken

    def getuserdeltanostate(self):
        """ Get a delta token with no state """

        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/users/delta?$select=displayName,givenName,surname,userPrincipalName,mail,manager&$deltaToken=latest"
        response = requests.get(request_string, headers=header)
        data = response.json()
        deltatoken = data['@odata.deltaLink']
        return deltatoken

    def getmanager(self, user):
        """ Gets the manager for a user """

        url_encoded = urllib.parse.quote_plus(f'{user}')
        request_string = f"{self.base_url}/users/{url_encoded}/manager"
        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        response = requests.get(request_string, headers=header)
        data = response.json()
        if "error" in data.keys():
            return "Notfound"
        else:
            return data['userPrincipalName']

    def assignmanager(self, user, manager):
        """ Creates an Enterprise app from the gallery given a template id """
        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        user = urllib.parse.quote_plus(f'{user}')
        mgrobj = f'{self.base_url}/users/{manager}'
        request_string = f"{self.base_url}/users/{user}/manager/$ref"
        request_body = json.dumps({
            "@odata.id": mgrobj
        })
        response = requests.put(
            request_string,
            data=request_body,
            headers=header)
        return response.status_code

    def checkmain(self, attrib):
        """ Get a delta token with no state """
        header = {
            "Content-type": "application/json",
            "Authorization": "Bearer " + self.access_token2
        }
        request_string = f"{self.base_url}/users/?$select={attrib}"
        response = requests.get(request_string, headers=header)
        data = response.json()
        userdata = data['value']
        next_url = ''
        while True:
            if '@odata.nextLink' in data:
                if data['@odata.nextLink'] == next_url:
                    break
                next_url = data['@odata.nextLink']
                next_data = requests.get(
                    next_url, headers=self.header_params_GMC).json()
                userdata += next_data['value']
                data = next_data
            else:
                break
        return userdata
