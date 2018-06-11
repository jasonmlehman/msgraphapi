# Microsoft GraphAPI interaction with Python

Office 365 administrator enthusiasts that live and breathe python may find this module useful.

# Installation

Install directly from this repo via pip.

* pip install git+https://github.com/jasonmlehman/msgraphapi.git

Alternatively, as I can't seem to get the above to work on Windows, you can clone the repo, change to the top-level directory (with the setup.py file) and use pip to install the local files in "editable" mode (-e).

* git clone https://github.com/jasonmlehman/msgraphapi.git
* cd msgraphapi
* pip install -e .

Create a creds.json file as dicated here, copied locally (i.e. /etc/creds.json)

https://github.com/jasonmlehman/msgraphapi/blob/master/msgraphapi/creds/creds.json

# How to use

Test config with the following command:

      listrolemembers -listrole TRUE -credpath /etc/creds.json

The above command should list all available office 365 roles within the tenant

To use within python to add a user to an office 365 role:

      from msgraphapi import msgraphapi

      credfile = "/etc/creds.json"
      userupn = "john.doe@somewhere.onmicrosoft.com"
      rolename = "Exchange Service Administrator"

      r = msgraphapi(credfile)

      # Get the directory objects ID
      userupnid = r.getupnid(userupn)
      roleid = r.getroleid(rolename)

      result = r.addusertorole(userupnid,roleid)
      print(result)

# Prerequisites

* Python 2.7
* requests module
* An Office 365 for business account

# Confusion with which RESTful endpoint to use

There is a lot of confusion as to which RESTful endpoint one should be using.  There is the Microsoft Graph API (graph.microsoft.com) and the Azure AD Graph API (graph.windows.net).  Microsofts roadmap calls for developers to use the Microsoft Graph API.  New functionality will be eventually be put into Microsoft Graph that may not be available with Azure AAD graph.  This module will show you how to use both API's.  There are some fundamental differences in which you interact with the API's.  Authentication tokens, headers, etc.

# Application Registration

There is a lot of confusion online about how to setup the application access for the Microsoft Graph API.  Some reads will have you create something within Azure Active directory and some will have you create something on apps.dev.microsoft.com.  I'll set the record straight: I ONLY created the application in Azure AD and didn't mess around with any app creation on apps.dev.microsoft.com.  This allowed me to authenticate using both API's.  Each time I tried to use the one created on apps.dev.microsoft.com I failed miserably.  I created my application using these steps:
  
  1)  Sign into Azure Active directory (portal.azure.com) as a global Admin (or simply have a global admin do this for you).
  2)  Select "Azure Active Directory" from the pane on the left
  3)  Select "Enterprise Applications"
  4)  Select "New Application"
  5)  Select "Application you are developing"
  6)  Select "Ok, take me to App Registrations to register my new application."
  7)  Select "New application registration"
  8)  In "name" put in a friendly name 
  9)  In the "Application Type" choose "Web App/API"
  10) In the "Sign in/url" just put in "http://localhost"
  11) Select "Create application"
  12) Select "Settings-Required Permissions"  You will have to give the permissions you will require depending on what you will do with       the API.  
  13) Select "Keys" and generate a secret
  14) Document application ID

# My personal use cases

I am a Global administrator for a large tenant.  I needed some automation within various locations of Office 365 and Azure AD.  A lot of this automation could have been done with an Azure AD powershell module.  But, the idea of powershell makes me sick so I decided to hit the graph API with python.  Some of the things I needed to do include:  
*    List all members of a given role within office 365 (i.e Get a list of all exchange service administrators)
*    Query a local active directory group for members and grant those members access to a role in office 365 (query a local AD group   *      called "AZ_Exchange_Admin" and grant all users within that group the "Exchange Service Administrator" role within office 365.          This would prevent me from having to micro manage that group.
*    Create Cloud only user accounts that aren't being federated with an existing directory
*    etc

Since my application would be managing office 365 roles I needed to grant the application global administrator privelages.  The process for doing this is:
  1)  Download "Active Directory PowerShell Module"
  2)  Open "Active Directory PowerShell Module" as administrator
  3)  Type "connect-msolservice"
  4)  When prompted enter global admin tenant permissions
  5)  Type "$sp = Get-MsolServicePrincipal -AppPrincipalId <App ID GUID>"  This is your application ID found from the application             created perviously.
  6)  Type "$role = Get-MsolRole -RoleName "Company Administrator"
  7)  Type "Add-MsolRoleMember -RoleObjectId $role.ObjectId -RoleMemberType ServicePrincipal -RoleMemberObjectId $sp.ObjectId"
  8)  To validate it's created you can type "Get-MsolRoleMember -RoleObjectId $role.ObjectId"

