#!/usr/bin/python
import requests
import json
import argparse
import sys

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

                # Get an access token from the graph.microsoft.com API
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


        def graphapirequest(self,request_string):
                header_params = {'Authorization': self.token_type + ' ' + self.access_token}
                response = requests.get(request_string, headers = header_params)
                data = response.json()
                return data

        def createuser(self,userupn,displayname,password,mailnickname):
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
                request_string = 'https://graph.windows.net/' + self.tenant + '/directoryRoles?api-version=beta'
                response = requests.get(request_string, headers=self.header_params)
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
                request_string = 'https://graph.microsoft.com/v1.0/users/' + upn
                response = requests.get(request_string, headers=self.header_params_GMC)
                data = response.json()
                return data['id']

        def listroles(self):
                request_string = 'https://graph.windows.net/' + self.tenant + '/directoryRoles?api-version=beta'
                response = requests.get(request_string, headers=self.header_params)
                data = response.json()
                roles = []
                for role in data['value']:
                        rolename = role['displayName']
                        roles.append(rolename)
                return roles

        def addusertorole(self, userguid, roleguid):
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

        def listrolemembers(self,roleguid):
                request_string = 'https://graph.windows.net/' + self.tenant + '/directoryRoles/' + roleguid + '/members/?api-version=beta'
                response = requests.get(request_string, headers=self.header_params)
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
