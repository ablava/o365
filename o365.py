#!/usr/bin/env python

"""
Simple script to manage Office365 users 
via Graph API. 

Usage: 
    python o365.py -f input.json -o output.csv

Options:
    -h --help
    -f --file	Input file (required)
    -o --out	Output file (required)

Environment specific script constants are stored in this 
config file: o365settings.py
    
Input:

Input file is expected to be in JSON format (e.g. input.json).
with these 10 required data fields:
{
    "useractions": [
        {
            "action": "create",
            "username": "testuserj",
            "newusername": "testuserj",
            "loginDisabled": "False",
            "UDCid": 1554943643675475475437,
            "givenName": "John",
            "fullName": "John The Testuser",
            "sn": "Testuser",
            "primO": "Biology",
            "userPassword": "initial password"
        }
    ] 
}
where action can be create/update/delete and newusername is same old one 
or a new value if renaming the user.

Note that this script only deletes users and it cannot purge/restore 
users from O365 RecycleBin where they stay for 30 days because
this functionality is not implemented in Graph API (yet).
You can use PowerShell to purge/restore deleted users with:
Remove-MsolUser -UserPrincipalName testuser@office365.com -RemoveFromRecycleBin
or to purge all of deleted users:
Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin -Force
and to restore a user:
$DelUser = Get-MsolUser -UserPrincipalName testu@o365.com -ReturnDeletedUsers
Restore-MsolUser -ObjectId $DelUser.ObjectId
    
Output:

Output file (e.g. output.csv) will have these fields:

action, username, result (ERROR/SUCCESS: reason)

Logging:

Script creates a detailed o365.log

All errors are also printed to stdout.

Author: A. Ablovatski
Email: ablovatskia@denison.edu
Date: 11/04/2016
"""

from __future__ import print_function
import time
import sys
import traceback
import json
import csv
import argparse
import logging
import httplib
import urllib


def main(argv):
    """This is the main body of the script"""
    
    # Setup the log file
    logging.basicConfig(
        filename='o365.log',level=logging.DEBUG, 
        format='%(asctime)s, %(levelname)s: %(message)s', 
        datefmt='%Y-%m-%d %H:%M:%S')

    # Get Azure creds and other constants from this settings file
    config_file = 'o365settings.py'
    
    if not readConfig(config_file):
        logging.error("unable to parse the settings file")
        sys.exit()
    
    # Parse script arguments
    parser = argparse.ArgumentParser()                                               

    parser.add_argument("--file", "-f", type=str, required=True, 
                        help="Input JSON file with user actions and params")
    parser.add_argument("--out", "-o", type=str, required=True, 
                        help="Output file with results of o365 user actions")

    try:
        args = parser.parse_args()
        
    except SystemExit:
        logging.error("required arguments missing - " \
                        "provide input and output file names")
        sys.exit()

    # Read input from json file
    in_file = args.file
    # Write output to csv file
    out_file = args.out
    
    try:
        f_in = open(in_file, 'rb')
        logging.info("opened input file: {0}".format(in_file))
        f_out = open(out_file, 'wb')
        logging.info("opened output file: {0}".format(out_file))
        reader = json.load(f_in)
        writer = csv.writer(f_out)
        writer.writerow(['action','username','result'])

        for row in reader["useractions"]:
            result = ''
            # Select what needs to be done
            if row["action"] == 'create':
                result = create(str(row["username"]), str(row["loginDisabled"]), 
                                str(row["UDCid"]), str(row["givenName"]), 
                                str(row["fullName"]), str(row["sn"]),  
                                str(row["primO"]), str(row["userPassword"]))
            elif row["action"] == 'update':
                result = update(str(row["username"]), str(row["newusername"]), 
                                str(row["loginDisabled"]), 
                                str(row["givenName"]), str(row["fullName"]), 
                                str(row["sn"]), str(row["primO"]))
            elif row["action"] == 'delete':
                 result = delete(str(row["username"]))
            else:
                print("ERROR: unrecognized action: {0}".format(row["action"]))
                logging.error("unrecognized action: {0}".format(row["action"]))
                result = "ERROR: Unrecognized action."
            
            # Write the result to the output csv file
            writer.writerow([row["action"], row["username"], result])
            
    except IOError:
        print("ERROR: Unable to open input/output file!")
        logging.critical("file not found: {0} or {1}".format(in_file, out_file))
        
    except Exception as e:
        traceb = sys.exc_info()[-1]
        stk = traceback.extract_tb(traceb, 1)
        fname = stk[0][3]
        print("ERROR: unknown error while processing line '{0}': " \
                "{1}".format(fname,e))
        logging.critical("unknown error while processing line '{0}': " \
                "{1}".format(fname,e))
        
    finally:
        f_in.close()
        logging.info("closed input file: {0}".format(in_file))
        f_out.close()
        logging.info("closed output file: {0}".format(out_file))
        
    return


def create(username, loginDisabled, UDCid, givenName, fullName, sn, ou, 
            userPassword):
    """This funtion adds users to O365"""
    
    # Check if any of the parameters are missing
    params = locals()
    
    for _item in params:
        if str(params[_item]) == "":
            print("ERROR: unable to create user {0} because {1} is missing " \
                    "a value".format(username, _item))
            logging.error("unable to create user {0} because {1} is missing " \
                            "a value".format(username, _item))
            result = "ERROR: Missing an expected input value for " + _item \
                        + " in input file."
            return result

    # Get the Graph API access_token and
    # Catch any MSFT login failures
    if not ACCESS_TOKEN:
        graphConnect()
        # Test again
        if not ACCESS_TOKEN:
            result = "ERROR: unable to authenticate to MSFT login service."
            return result
    
    # Grab the access_token
    access_token = ACCESS_TOKEN
    
    # Do a quick check if the user already exists
    upn = username + "@" + O365DOMAIN

    if findUser(upn):
        print("ERROR: cannot create user - user already exists: {0}" \
                .format(username))
        logging.error("cannot create user - user already exists: {0}" \
                .format(username))
        result = "ERROR: username already taken!"
        return result
    
    try:
        # Set the bearer auth header
        headers = {
            'Authorization': 'Bearer ' + access_token,
            'Content-Type': 'application/json; charset=utf-8'
        }
        
        # Set the required params
        params = urllib.urlencode({
            'api-version': API_VERSION,
        })

        # Create body of the request                
        body = {
            "userPrincipalName": upn,
            "accountEnabled": "true",
            "givenName": givenName,
            "displayName": fullName,
            "surname": sn,
            "mailNickname": username,
            "department": ou,
            "immutableId": UDCid,
            "passwordProfile": {"password": userPassword, 
                                "forceChangePasswordNextLogin": "false"},
            "passwordPolicies": "DisablePasswordExpiration",
            "usageLocation": "US"
        }
        
        data = json.dumps(body)
        
        # Connect to o365
        conn = httplib.HTTPSConnection('graph.windows.net')
        conn.request("POST", "/" + O365DOMAIN + "/users?" 
                        + params, data, headers)
        response = conn.getresponse()
        
        if response.status != 201:
            # User was not created
            logging.error("user could not be created in o365: {0}" \
                        .format(username))
            print("ERROR: User {0} could not be added to o365" \
                        .format(username))
            result = "ERROR: user could not be created in o365."
        else:
            # Close earlier connection
            conn.close()
            
            # Wait 5s for user to be fully created
            #time.sleep(5)
            
            # Determine licenses based on user type
            userType = getUserType(username)
            
            # Assign correct licenses for Stu/Emp users
            # Disable MCOSTANDARD/EXCHANGE_S_STANDARD plans
            if userType == "STU":
                licenses = [{"disabledPlans":DISABLEDPLANS, "skuId":STULICENSE}]
            else:
                licenses = [{"disabledPlans":DISABLEDPLANS, "skuId":EMPLICENSE}]
            
            # Now try assigning licenses
            body = {
                "addLicenses": licenses,
                "removeLicenses": []
            }
            
            data = json.dumps(body)
            
            conn = httplib.HTTPSConnection('graph.windows.net')
            conn.request("POST", "/" + O365DOMAIN + "/users/" + upn 
                            + "/assignLicense?" + params, data, headers)
            response = conn.getresponse()
            
            if response.status != 200:
                # User was created with no licenses
                logging.error("user did not get licenses in o365: {0}" \
                        .format(username))
                print("ERROR: User {0} did not get licenses in o365" \
                        .format(username))
                result = "SUCCESS: user added but with no licenses in o365."
            else:
                # Log user creation
                logging.info("user added to o365: {0}".format(username))
                print("SUCCESS: User {0} added to o365".format(username))
                result = "SUCCESS: user was created in o365."
            
        conn.close()
        
    except Exception as e:
        print("ERROR: Could not add user to o365: {0}".format(e))
        logging.error("o365 add failed for user: {0}: {1}".format(username,e))
        result = "ERROR: could not create o365 user."
        return result
    
    return result


def update(username, newusername, loginDisabled, givenName, fullName, sn, ou):
    """This function updates user attributes, 
    blocks and renames users if needed"""
    
    # Note: we can't change UDCid - it is an ImmutableId in O365!

    # Check if any of the arguments are missing
    params = locals()
    
    for _item in params:
        if str(params[_item]) == "":
            print("ERROR: unable to update user {0} because {1} is missing " \
                    "a value".format(username, _item))
            logging.error("unable to update user {0} because {1} is missing " \
                            "a value".format(username, _item))
            result = "ERROR: Missing an expected input value for " \
                        + _item + " in input file."
            return result

    # Get the Graph API access_token and
    # Catch any MSFT login failures
    if not ACCESS_TOKEN:
        graphConnect()
        if not ACCESS_TOKEN:
            result = "ERROR: unable to authenticate to MSFT login service."
            return result
    
    access_token = ACCESS_TOKEN
    
    # Do a quick check if the user already exists
    upn = username + "@" + O365DOMAIN
    
    if not findUser(upn):
        print("ERROR: user does not exist in o365: {0}".format(username))
        logging.error("user does not exist in o365: {0}".format(username))
        result = "ERROR: user could not be found in o365!"
        return result
    
    # rename if new username is diferent
    # and does not already exist
    if username != newusername:
        newupn = newusername + "@" + O365DOMAIN
        
        # Check if the new user name already exists
        if findUser(newupn):
            print("ERROR: cannot rename user - user already exists: {0}" \
                    .format(newusername))
            logging.error("cannot rename user - user already exists: {0}" \
                            .format(newusername))
            result = "ERROR: username already taken!"
            return result
            
    # Rename, update attributes or disable
    try:
        # Set the bearer auth header
        headers = {
            'Authorization': 'Bearer ' + access_token,
            'Content-Type': 'application/json; charset=utf-8'
        }
        
        # Set the required params
        params = urllib.urlencode({
            'api-version': API_VERSION,
        })
        
        # Flip the loginDisabled value
        if loginDisabled == "True":
            accountEnabled = "False"
        else:
            accountEnabled = "True"
        
        body = {
            "userPrincipalName": newupn,
            "accountEnabled": accountEnabled,
            "givenName": givenName,
            "displayName": fullName,
            "surname": sn,
            "mailNickname": newusername,
            "department": ou
        }
 
        data = json.dumps(body)
        
        conn = httplib.HTTPSConnection('graph.windows.net')
        conn.request("PATCH", "/" + O365DOMAIN + "/users/" + upn + "?" 
                        + params, data, headers)
        response = conn.getresponse()
        
        if response.status != 204:
            logging.error("user was not updated in o365: {0}" \
                    .format(username))
            print("ERROR: User {0} was not updated in o365" \
                    .format(username))
            result = "ERROR: could not update user in o365."
        else:
            logging.info("user updated o365: {0}".format(username))
            print("SUCCESS: User {0} updated in o365".format(username))
            result = "SUCCESS: user was updated in o365."
            
        conn.close()
        
    except Exception as e:
        print("ERROR: Could not update user in o365: {0}".format(e))
        logging.error("o365 update failed for: {0}: {1}".format(username,e))
        result = "ERROR: Could not update o365 user."
        return result
    
    return result


def delete(username):
    """This function deletes a user from O365"""

    # Check if the argument is missing
    if str(username) == "":
        print("ERROR: unable to delete user because username argument " \
                "is missing a value")
        logging.error("unable to delete user because username argument " \
                        "is missing a value")
        result = "ERROR: Missing an expected input value for username " \
                    "in input file."
        return result

    # Get the Graph API access_token and
    # Catch any MSFT login failures
    if not ACCESS_TOKEN:
        graphConnect()
        if not ACCESS_TOKEN:
            result = "ERROR: unable to authenticate to MSFT login service."
            return result
    
    access_token = ACCESS_TOKEN
    
    # Do a quick check if the user exists
    upn = username + "@" + O365DOMAIN

    if not findUser(upn):
        print("ERROR: user does not exist in o365: {0}".format(username))
        logging.error("user does not exist in o365: {0}".format(username))
        result = "ERROR: user could not be found in o365!"
        return result
        
    # Delete the user if all is OK
    try:
         # Build bearer auth header
        headers = {
            'Authorization': 'Bearer ' + access_token,
            'Content-Type': 'application/json'
        }
    
        params = urllib.urlencode({
            'api-version': API_VERSION,
        })
    
        # Connect to Graph API
        conn = httplib.HTTPSConnection('graph.windows.net')
        conn.request("DELETE", "/" + O365DOMAIN + "/users/" + upn + "?" 
                        + params, "", headers)
        response = conn.getresponse()

        if response.status != 204:
            logging.error("user was not deleted in o365: {0}" \
                    .format(username))
            print("ERROR: User {0} was not deleted in o365" \
                    .format(username))
            result = "ERROR: could not delete user in o365."
        else:
            logging.info("user deleted in o365: {0}".format(username))
            print("SUCCESS: User {0} deleted in o365".format(username))
            result = "SUCCESS: user deleted in o365."
            
        conn.close()

    except Exception as e:
        print("ERROR: unknown error while deleting user: {0}".format(e))
        logging.error("unknown error while deleting user {0}: {1}" \
                        .format(username,e))
        result = "ERROR: Could not delete o365 user."
    
    return result


def readConfig(config_file):
    """Function to import the config file"""
    
    if config_file[-3:] == ".py":
        config_file = config_file[:-3]
    o365settings = __import__(config_file, globals(), locals(), [])
    
    # Read settings and set globals
    try: 
        global API_VERSION
        global CLIENT_ID
        global CLIENT_KEY
        global O365DOMAIN
        global STUPATTERN
        global STULICENSE
        global EMPLICENSE
        global DISABLEDPLANS

        API_VERSION = o365settings.API_VERSION
        CLIENT_ID = o365settings.CLIENT_ID
        CLIENT_KEY = o365settings.CLIENT_KEY
        O365DOMAIN = o365settings.O365DOMAIN
        STUPATTERN = o365settings.STUPATTERN
        STULICENSE = o365settings.STULICENSE
        EMPLICENSE = o365settings.EMPLICENSE
        DISABLEDPLANS = o365settings.DISABLEDPLANS
        
        global ACCESS_TOKEN
        ACCESS_TOKEN = None

    except Exception as e:
        logging.error("unable to parse settings file")
        print("ERROR: unable to parse the settings file: {0}".format(e))
        return False
        
    return True


def getUserType(username):
    """ Function to determine the type of a user"""

    if STUPATTERN in username:
        userType = "STU"
    else:
        userType = "EMP"

    return userType


def graphConnect():
    """This function gets auth token for Graph API"""

    # Use global ACCESS_TOKEN variable
    global ACCESS_TOKEN
    
    # Set the connecton parameters
    params = urllib.urlencode({
        'api-version': API_VERSION,
    })

    # Create the request body
    body = urllib.urlencode({
        "redirect_uri" : "http://127.0.0.1:8000",
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_KEY,
        "resource": "https://graph.windows.net/"
    })

    # Create the request headers
    headers = {
    }

    try:
        # Open a connection to the MSFT login service
        conn = httplib.HTTPSConnection('login.windows.net')
        conn.request("POST", "/" + O365DOMAIN + "/oauth2/token?" + params, body, 
                        headers)
        response = conn.getresponse()

        # Get the auth token
        if response.status == 200:
            data = response.read()
            jsondata = json.loads(data)
            ACCESS_TOKEN = jsondata['access_token']
        else:
            print("ERROR: MSFT login service did not respond correctly")
            logging.error("MSFT login service returned: {0}" \
                            .format(response.status))
        conn.close()
        
    except Exception as e:
        print("ERROR: Could not connect to MSFT login service: {0}".format(e))
        logging.error("problem connecting to MSFT login service: {0}".format(e))
     
    return


def findUser(upn):
    """Do a quick check if the user already exists"""
    
    # Grab the access_token to Graph API
    access_token = ACCESS_TOKEN
    
    # Build bearer auth header
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }
    
    params = urllib.urlencode({
        'api-version': API_VERSION,
    })

    try:
        # Connect to Graph API
        conn = httplib.HTTPSConnection('graph.windows.net')
        conn.request("GET", "/" + O365DOMAIN + "/users/" + upn + "?" 
                        + params, "", headers)
        response = conn.getresponse()
        conn.close()
        
        # Check if the user does not exist
        if response.status != 200:
            logging.info("user {0} does not exist in o365".format(upn))
            return False
        
    except Exception as e:
        print("ERROR: problem with user search in O365: {0}".format(e))
        logging.error("problem searching for {0} in O365: {1}".format(upn,e))
        
    return True


if __name__ == "__main__":
    main(sys.argv)
