from msgraphapi import msgraphapi
import argparse
import sys

# Lists all the members of an Office 365 Role
# Usage example: python listrolemembers.py -role "Company administrator" -environment test
# To get all roles available to query use: python listrolemembers.py -listrole TRUE -environment test

parser = argparse.ArgumentParser()
parser.add_argument('-role', '--role',
                    help='The Office 365 Role to query')
parser.add_argument('-environment', '--environment',
                    help='select environment: test/prod')
parser.add_argument('-listrole', '--listrole',
                    help='If argument is set will list all available office 365 roles')

args = parser.parse_args()
role = args.role
env = args.environment
rolelist = args.listrole

if rolelist:
        r = msgraphapi(environment=env)
        result = r.listroles()
        for roles in result:
                print roles
        sys.exit(1)

if role == None:
        print("Please specify role using -role parameter")
        sys.exit(1)

if env == None:
        print("Please specify environment (test/prod) using -environment parameter")
        sys.exit(1)

r = msgraphapi(environment=env)

# Need to get the ID of the directory object for the role

roleid = r.getroleid(role)

#  Now that we have directory objects we can make the request

result = r.listrolemembers(roleid)
for roles in result:
        print roles
