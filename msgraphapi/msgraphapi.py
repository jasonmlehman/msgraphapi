#!/usr/bin/python
from __future__ import division
import requests
import json
import argparse
import sys
import datetime

class msgraphapi:

        #constants

        def __init__(self, credpath):
		# credpath is the path to a json file containing tenant credentials
                creds = json.load(open(credpath))
		self.clientid = creds["clientid"]
                self.client_secret = creds["clientsecret"]
                self.login_url = creds["loginurl"]
                self.tenant = creds["tenant"]

                # Get an access token for the graph.windows.net API
                self.bodyvals = {'client_id': self.clientid,
                    'client_secret': self.client_secret,
                    'grant_type': 'client_credentials'}

                self.request_url = self.login_url + self.tenant + '/oauth2/token?api-version=beta'
                self.token_response = requests.post(self.request_url, data=self.bodyvals)
                self.access_token = self.token_response.json().get('access_token')
                self.token_type = self.token_response.json().get('token_type')
                if self.access_token is None or self.token_type is None:
                        print("ERROR: Couldn't get access token")
                        sys.exit(1)
                self.header_params = {'Authorization': self.token_type + ' ' + self.access_token}

                # Get an access token for the graph.microsoft.com API
                self.bodyvals = dict(
                        client_id=self.clientid,
                        client_secret=self.client_secret,
                        grant_type='client_credentials',
                        resource='https://graph.microsoft.com',
                )

                self.header_params2 = {
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'Accept': 'application/json'
                }

                self.request_url = self.login_url + self.tenant + '/oauth2/token'
                self.token_response = requests.post(self.request_url, data=self.bodyvals, headers=self.header_params2)

                self.access_token2 = self.token_response.json().get('access_token')
                self.token_type = self.token_response.json().get('token_type')
                if self.access_token is None or self.token_type is None:
                        print("ERROR: Couldn't get access token")
                        sys.exit(1)
                self.header_params_GMC = {'Authorization': self.token_type + ' ' + self.access_token2}

		# Get an access token for the Office 365 Management API
                self.bodyvals = dict(
                        client_id=self.clientid,
                        client_secret=self.client_secret,
                        grant_type='client_credentials',
                        resource='https://manage.office.com',
                )

                self.header_params3 = {
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'Accept': 'application/json'
                }

                self.request_url = self.login_url + self.tenant + '/oauth2/token'
                self.token_response = requests.post(self.request_url, data=self.bodyvals, headers=self.header_params2)

                self.access_token3 = self.token_response.json().get('access_token')
                self.token_type = self.token_response.json().get('token_type')
                if self.access_token3 is None or self.token_type is None:
                        print("ERROR: Couldn't get access token")
                        sys.exit(1)
                self.header_params_MOC = {'Authorization': self.token_type + ' ' + self.access_token3}

        def graphapirequest(self,request_string):
                header_params = {'Authorization': self.token_type + ' ' + self.access_token}
                response = requests.get(request_string, headers = header_params)
                data = response.json()
                return data

        def createuser(self,userupn,displayname,password,mailnickname):
		# Creates a user in the tenant
                header_params = {
                        'Authorization': self.token_type + ' ' + self.access_token,
                        'Content-Type': 'application/json'
                }
                request_body = {'accountEnabled': 'true',
                        'displayName': displayname,
                        'passwordProfile':
                                {
                                        'password': password,
                                        'forceChangePasswordNextLogin': 'true'
                                },
                        'mailNickname': mailnickname,
                        'userPrincipalName': userupn}

                request_string = 'https://graph.windows.net/' + self.tenant + '/users?api-version=beta&'
                response = requests.post(request_string, data=json.dumps(request_body), headers=header_params)
                data = response.json()
                return data

        def getroleid(self,rolename):
		# Gets the GUID given a role name
                request_string = 'https://graph.windows.net/' + self.tenant + '/directoryRoles?api-version=1.6'
		while True:
			try:
		                response = requests.get(request_string, headers=self.header_params)
				break
			except Exception:
					pass
                data = response.json()
                objectid = ""
                for role in data['value']:
                        if role['displayName'].lower() == rolename.lower():
                                dispname=role['displayName']
                                objectid=role['objectId']
                if objectid:
                        return objectid
                else:
                        return "Role not found"

        def getupnid(self, upn):
		# Gets the guid of a user given the UPN name
                request_string = 'https://graph.microsoft.com/v1.0/users/' + upn
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                return data['id']

        def getb2binfo(self, guid):
                request_string = 'https://graph.microsoft.com/v1.0/users/' + guid
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                return data

	def getusers(self):
		request_string = 'https://graph.microsoft.com/v1.0/users/'
		response = requests.get(request_string, headers=self.header_params_GMC)
		data = response.json()
		users = []
		for user in data['value']:
			if user['mail'] != "None":
				users.append(user['mail'])
		return json.dumps(users, indent=4, sort_keys=True)

        def getmailusers(self):
               	request_string = "https://graph.microsoft.com/beta/reports/getMailboxUsageDetail(period='D7')?$format=application/json"
               	response = requests.get(request_string, headers=self.header_params_GMC)
		data = response.json()
		users = []
		for user in data['value']:
			username = user['userPrincipalName']
			users.append(username)
                return json.dumps(data, indent=4, sort_keys=True)

        def listroles(self):
		# Lists all roles in the tenant
                request_string = 'https://graph.windows.net/' + self.tenant + '/directoryRoles?api-version=beta'
                response = requests.get(request_string, headers=self.header_params)
                data = response.json()
                roles = []
                for role in data['value']:
                        rolename = role['displayName']
                        roles.append(rolename)
                return roles

        def addusertorole(self, userguid, roleguid):
		# Adds a user to a role given the userguid and role guid
                header = {
                        "Content-type": "application/json",
                        "Authorization": "Bearer " + self.access_token
                }

                request_string = 'https://graph.windows.net/' + self.tenant + '/directoryRoles/' + roleguid + '/$links/members?api-version=beta'
                request_body = json.dumps({
                        "url": "https://graph.windows.net/" + self.tenant + "/directoryObjects/" + userguid
                })
                response = requests.post(request_string, data=request_body, headers=header)

                try:
                        data = response.json()
                        error = data["odata.error"]["message"]["value"]
                        return error
                except:
                        return "Success"

        def removeuserfromrole(self, userguid, roleguid):
		# Removes a user from a role
                header = {
                        "Content-type": "application/json",
                        "Authorization": "Bearer " + self.access_token
                }

                request_string = 'https://graph.windows.net/' + self.tenant + '/directoryRoles/' + roleguid + '/$links/members/' + userguid + '?api-version=beta'
		response = requests.delete(request_string, headers=header)
                try:
                        data = response.json()
                        error = data["odata.error"]["message"]["value"]
                        return error
                except:
                        return "Success"

        def listrolemembers(self,roleguid):
		# Lists the members assgined to a role guid
                request_string = 'https://graph.windows.net/' + self.tenant + '/directoryRoles/' + roleguid + '/members/?api-version=1.6'
		while True:
			try:
		                response = requests.get(request_string, headers=self.header_params)
				break
                        except Exception:
                                pass
                data = response.json()
                members = []
                for member in data['value']:
                        if member['objectType'] == "User":
                                membername = member['userPrincipalName']
                                members.append(membername)
                        else:
                                membername = member['displayName']
                                members.append(membername)
                return members

        def activitylog(self, numofdays):
                # Access audit data
                # Who can access?  Users in Security Admin role or Security Reader Role, or Global Admins
                startdate = datetime.date.strftime(datetime.date.today() - datetime.timedelta(days=numofdays), '%Y-%m-%d')
                request_string = 'https://graph.windows.net/' + self.tenant + '/activities/audit?api-version=beta&$filter=activityDate gt ' + startdate
                response = requests.get(request_string, headers=self.header_params)
                data = response.json()
                return json.dumps(data)

        def auditlogbyid(self, directoryid):
                # Provides (or gets) a specific Azure Active Directory audit log item.
                # Permissions needed
                #       Dellegated (work or school account): AuditLog.Read.All
                #       Application:    AuditLog.Read.All

                request_string = 'https://graph.microsoft.com/beta/auditLogs/directoryAudits/' + directoryid
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                return json.dumps(data)

        def auditlog(self, numofdays):
                # Provides the list of audit logs generated by Azure Active Directory
                # Permissions needed
                #       Dellegated (work or school account): AuditLog.Read.All
                #       Application:    AuditLog.Read.All

                startdate = datetime.date.strftime(datetime.date.today() - datetime.timedelta(days=numofdays), '%Y-%m-%d')
                request_string = 'https://graph.microsoft.com/beta/auditLogs/directoryAudits?&$filter=activityDateTime ge ' + startdate
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()

                return json.dumps(data, indent=4, sort_keys=True)

        def listsubscriptions(self, publisherid):
		request_string = 'https://manage.office.com/api/v1.0/' + self.tenant + '/activity/feed/subscriptions/list?PublisherIdentifier=' + publisherid

                response = requests.get(request_string, headers=self.header_params_MOC)
                data = response.json()
                return json.dumps(data)

	def createsubscription(self, contenttype):
		# Contentype must be Audit.SharePoint, Audit.AzureActiveDirectory, Audit.Exchange, Audit.General, or DLP.All

		request_string = 'https://manage.office.com/api/v1.0/' + self.tenant + '/activity/feed/subscriptions/start?contentType=' + contenttype

                response = requests.post(request_string, headers=self.header_params_MOC)
                data = response.json()
                return json.dumps(data)

	def getriskyevents(self, numofdays='90'):
		# Gest the last 90 days of risky events

                request_string = 'https://graph.microsoft.com/v1.0/identityRiskEvents/'
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                return json.dumps(data, indent=4, sort_keys=True)

        def listrolesv2(self):
		# List Office 365 roles

                request_string = 'https://graph.microsoft.com/beta/directoryRoles'
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
	        roles = [x['displayName'] for x in data['value']]
                return json.dumps(roles)

        def getroletemplates(self):
                request_string = 'https://graph.microsoft.com/beta/directoryRoles'
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                return json.dumps(data, indent=4, sort_keys=True)

        def activaterole(self,roleid):
		# Not all roles are available via the API by default, this activates the ones that aren't
		header = {
			"Content-type": "application/json",
			"Authorization": "Bearer " + self.access_token2
		}
                request_string = 'https://graph.microsoft.com/beta/directoryRoles'
		request_body = json.dumps({
			"roleTemplateId": roleid
		})
                response = requests.post(request_string, data=request_body,headers=header)
                data = response.json()
                return json.dumps(data, indent=4, sort_keys=True)

        def addmembertorole(self,upn,roleguid):
		# Adds a member to a role

                header = {
                        "Content-type": "application/json",
                        "Authorization": "Bearer " + self.access_token2
                }
                request_string = 'https://graph.microsoft.com/beta/directoryRoles/%s/members/$ref' % roleguid
		print self.getroleid("teams service administrator")
		print request_string
		userguid = self.getupnid(upn)
		print userguid
                request_body = json.dumps({
                        "id": userguid
                })
                response = requests.post(request_string, data=request_body,headers=header)
                data = response.json()
                return json.dumps(data, indent=4, sort_keys=True)

        def updatememberuser(self,guid):
		# This changes a guest user to a member of the directory

                header = {
                        "Content-type": "application/json",
                        "Authorization": "Bearer " + self.access_token2
                }
                request_string = 'https://graph.microsoft.com/v1.0/users/%s' % guid
                print(request_string)
                request_body = json.dumps({
                        "userType": "Member"
                })
                response = requests.patch(request_string, data=request_body,headers=header)
                return response 

        def getauditdata(self, logtype, folderpath=None, publisherid='self.tenant'):
		# Gets the last 24 hours worth of content
		# Contentype must be Audit.SharePoint, Audit.AzureActiveDirectory, Audit.Exchange, Audit.General, or DLP.All

                request_string = 'https://manage.office.com/api/v1.0/' + self.tenant + '/activity/feed/subscriptions/content?contentType=' + logtype + '&PublisherIdentifier=' + publisherid

                page0 = requests.get(request_string, headers=self.header_params_MOC)
                data = page0.json()
		
		# Get all the content Blob URI's
		blobdata = []
		bloburis = [blob['contentUri'] for blob in data]
		for bloburi in bloburis:
			blobr = requests.get(bloburi, headers=self.header_params_MOC)
			data = blobr.json()
			blobdata.append(data)

		# check response header to see if there is another page of data
		h = page0.headers
		print(h)
		if 'NextPageUri' in h:
			nextpage = h['NextPageUri']
			request_string = nextpage
			while True:
				print request_string
				pagex = requests.get(request_string, headers=self.header_params_MOC)
				data = pagex.json()
				bloburis = [blob['contentUri'] for blob in data]				
				for bloburi in bloburis:
					blobr = requests.get(bloburi, headers=self.header_params_MOC)
					data = blobr.json()
					blobdata.append(data)
				try:
					h = pagex.headers
					nextpage = h['NextPageUri']
					request_string = nextpage
				except:
					break


                if folderpath != None:
                        now = datetime.datetime.now()
                        date = now.strftime("%d-%m-%Y")

		return json.dumps(blobdata, indent=4, sort_keys=True)

        def skuinuse(self,skuid):
		# For licensing reporting, gets the number of units used for a skuid

                request_string = "https://graph.microsoft.com/v1.0/subscribedSkus"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
		for skus in data['value']:
			if skus['id'] == skuid:
				total = skus['prepaidUnits']['enabled']
				inuse = skus['consumedUnits']
				percentinuse = (inuse/total * 100)
		return inuse,total,percentinuse

	def getsecurescore(self):
		# Gets secure score stuff

		request_string = "https://graph.microsoft.com/stagingBeta/reports/getTenantSecureScores(period=1)/content"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                return json.dumps(data, indent=4, sort_keys=True)

        def getsignins(self):
		# Gets signins from the audit logs

                request_string = "https://graph.microsoft.com/beta/auditLogs/signIns"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
		users = []
		for user in data['value']:
			users.append(user['userPrincipalName'])
		return users

        def getdirobject(self, directoryid):
                request_string = "https://graph.microsoft.com/v1.0/directoryObjects/" + directoryid
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
		return data

        def getgroupnamefromid(self, directoryid):
                request_string = "https://graph.microsoft.com/v1.0/directoryObjects/" + directoryid
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()

		obj = self.getdirobject(directoryid)
		groupname = obj["displayName"]

                return groupname

        def listmembersofgroup(self, directoryid):
		# Lists the members of a group

                request_string = "https://graph.microsoft.com/v1.0/groups/" + directoryid + "/members"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
		obj = self.getdirobject(directoryid)
		groupname = obj["displayName"]
		
		members = []
		for member in data['value']:
			if member["@odata.type"] == "#microsoft.graph.user":
				upn = member['userPrincipalName']
				members.append(upn)
		return groupname, members

        def listguidofgroup(self, directoryid):
		# Gets all the guids of the members of a group

                request_string = "https://graph.microsoft.com/v1.0/groups/" + directoryid + "/members"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                memberguids = []
                for member in data['value']:
			if member["@odata.type"] == "#microsoft.graph.user":
				guid = member['id']
				memberguids.append(guid)
                return memberguids

        def listmembers(self, directoryid):
		# Gets the UPN's of users of a group

                request_string = "https://graph.microsoft.com/v1.0/groups/" + directoryid + "/members"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()

                member_list = data['value']
                next_url = ''
                while True:
                        if '@odata.nextLink' in data:
                                if data['@odata.nextLink'] == next_url:
                                        break
                                next_url = data['@odata.nextLink']
                                next_data = requests.get(next_url, headers=self.header_params_GMC).json()
                                member_list += next_data['value']
                                data = next_data
                        else:
                                break
		membersupn = [x['userPrincipalName'] for x in member_list if 'userPrincipalName' in x]
                return membersupn

        def getgroupdeletedate(self):
		# Gets all Office 365 groups deleted in the last 30 days

                request_string = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group"
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
			groupdict.update( {group['displayName'] : pdtstr} )
			sorted_g = sorted(groupdict.items(), key=lambda x: x[1])
		return json.dumps(sorted_g, indent=4)

	def geto365groups(self):
		# Gets the GUIDs of all Office 365 groups

                request_string = "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                groups = []
                for group in data['value']:
			groups.append(group['id'])
		return groups

        def geto365groupowner(self, directoryid):
		# Gets all office 365 groups without owners

                request_string = "https://graph.microsoft.com/v1.0/groups/" + directoryid + "/owners"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                owners = []
                for owner in data['value']:
                       owners.append(owner['userPrincipalName'])
                return owners

        def inviteuser(self,usermail,url):
		# Invites a guest to a tenant without sending an email
		# Change sendInvitationMessage to True to trigger an email

                header = {
                        "Content-type": "application/json",
                        "Authorization": "Bearer " + self.access_token2
                }
                request_string = 'https://graph.microsoft.com/v1.0/invitations'
                request_body = json.dumps({
                        "invitedUserEmailAddress": usermail,
			"inviteRedirectUrl": url,
			"sendInvitationMessage": False
                })
                response = requests.post(request_string, data=request_body,headers=header)
                data = response.json()
                return json.dumps(data, indent=4, sort_keys=True)


        def addmembertogroup(self,upn,groupguid):
		# Adds a user to a gropu

                header = {
                        "Content-type": "application/json",
                        "Authorization": "Bearer " + self.access_token2
                }
                request_string = 'https://graph.microsoft.com/v1.0/groups/' + groupguid + 'members/$ref'
                userguid = self.getupnid(upn)
                print userguid
                request_body = json.dumps({
                        "id": userguid
                })
                response = requests.post(request_string, data=request_body,headers=header)
                data = response.json()
                return json.dumps(data, indent=4, sort_keys=True)

        def addguidtogroup(self,userguid,groupguid):
		# Adds a user guid to a group guid

                header = {
                        "Content-type": "application/json",
                        "Authorization": "Bearer " + self.access_token2
                }
                request_string = 'https://graph.microsoft.com/v1.0/groups/' + groupguid + '/members/$ref'
                request_body = json.dumps({
                        "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + userguid
                })
                response = requests.post(request_string, data=request_body,headers=header)
                return response

	def getguestusers(self):
		# Gets all guest users in a tenant

                request_string = "https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Guest'"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
		return data

	def deleteuser(self, userid):
		# Deletes a user from a tenant

                request_string = "https://graph.microsoft.com/v1.0/users/%s/" % userid
                response = requests.delete(request_string, headers=self.header_params_GMC)
                return response

        def getdistgroups(self):
		# Gets all mail enabled groups
		
                request_string = "https://graph.microsoft.com/v1.0/groups"
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                data = [x['displayName'] for x in data['value'] if x['mailEnabled'] is True]
                return json.dumps(data, indent=4, sort_keys=True)
