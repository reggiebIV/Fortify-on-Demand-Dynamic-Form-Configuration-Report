import sys
import requests
import json
import math
import xlsxwriter
import time
import datetime
import logging
import os

#Check if there is a log directory in the current working directory, and create one if there is not
if not os.path.exists('log'):
    os.makedirs('log')

#This instantiates a logger that will be used to track any applications and dynamic forms that don't correctly populate in FoD
logger = logging.getLogger('FodDynamicUpdate')
hdlr = logging.FileHandler('log/FodDynamicUpdate.log')
formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
logger.setLevel(logging.INFO)
logger.addHandler(hdlr)

#Check if there is a log directory in the current working directory, and create one if there is not
if not os.path.exists('report'):
    os.makedirs('report')

#Within the log directory, check for a file named FodImport.log, and create one if there is not
workbook = xlsxwriter.Workbook('report/DynamicConfigurationReport2.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0,0, 'Application')
worksheet.write(0,1, 'Site URL')
worksheet.write(0,2, 'Authorization Type')    
worksheet.write(0,3, 'Primary User Name')
worksheet.write(0,4, 'Secondary User Name')
worksheet.write(0,5, 'Other User Name')    
worksheet.write(0,6, 'Site Facing')
worksheet.write(0,7, 'Entitlement Frequency Type')
worksheet.write(0,8, 'Time Zone')  

def getAllReleases(apiKey, apiSecret):
    bearerToken = GetToken(apiKey, apiSecret)
    getReleasesUrl = "https://api.ams.fortify.com/api/v3/releases/"
    allReleases = []

    headers = {
        'authorization': "Bearer " + bearerToken,
        'Accept': "application/json"
    }
    response        = requests.request("GET", getReleasesUrl, headers=headers)
    fullResponse    = json.loads(response.text)
    allReleasesObj  = fullResponse['items']
    numReleaseInObj = len(allReleasesObj)
    count = fullResponse['totalCount']

    numLoops = math.ceil(count / 50)

    for loop in range(0, numLoops):
        if loop == 0:
            for n in range(0, numReleaseInObj):
                thisRelease             = {}
                thisRelease['id']       = allReleasesObj[n]['releaseId']
                thisRelease['name']     = allReleasesObj[n]['releaseName']
                thisRelease['appId']    = allReleasesObj[n]['applicationId']
                thisRelease['appName']  = allReleasesObj[n]['applicationName']
                allReleases.append(thisRelease)
        else:
            # This increases the offset each time to get the next batch of users from the API
            offset          = loop * 50
            getReleasesUrl  = "https://api.ams.fortify.com/api/v3/releases?offset=" + str(offset)
            response        = requests.request("GET", getReleasesUrl, headers=headers)
            fullResponse    = json.loads(response.text)
            allReleasesObj  = fullResponse['items']
            numReleaseInObj = len(allReleasesObj)
            
            for z in range(0, numReleaseInObj):
                thisRelease             = {}
                thisRelease['id']       = allReleasesObj[z]['releaseId']
                thisRelease['name']     = allReleasesObj[z]['releaseName']
                thisRelease['appId']    = allReleasesObj[z]['applicationId']
                thisRelease['appName']  = allReleasesObj[z]['applicationName']
                allReleases.append(thisRelease)

    parseReleaseData(allReleases, bearerToken)
    
def parseReleaseData(allReleases, bearerToken):
    numberReleases = len(allReleases)
    rowCount = 1
    for i in range(0,numberReleases):
        thisRelease = allReleases[i]
        releaseId = thisRelease['id']
        if releaseId != 0:
            dynamicConfig = getDynamicConfig(releaseId, bearerToken)
            if dynamicConfig != None:
                print(dynamicConfig['dynamicSiteURL'])
                if dynamicConfig['dynamicSiteURL']:
                    try:
                        percentComplete = str(round((i/numberReleases)*100))
                        print(percentComplete + "%")
                        generateReportRow(dynamicConfig, thisRelease['appName'], rowCount)
                        rowCount+=1
                    except:
                        print("Error on ID " + str(releaseId))
                        ts = time.time()
                        messageForLog = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S') + " ERROR Release: " + str(releaseId)
                        logger.info(messageForLog)
                
    workbook.close()
    
def getReleaseId(appId, bearerToken):
    # When you create an application via the API, you are required to create a first release as well. But the API only returns the application ID
    # This method gets the ID of the release that you created so that you can use it to fill out the dynamic form (which is associated with releases
    # not applications)
    appIdString = str(appId)
    releaseDataUrl = "https://api.ams.fortify.com/api/v3/applications/" + appIdString + "/releases"
    headers = {
        'authorization': "Bearer " + bearerToken,
        'Accept': "application/json"
    }
    response = requests.request("GET", releaseDataUrl, headers=headers)
    fullResponse = json.loads(response.text)
    print(fullResponse['totalCount'])
    if int(fullResponse['totalCount']) > 0:
        releaseId = fullResponse['items'][0]['releaseId']
    else:
        releaseId = 0

    return releaseId

def GetToken(apiKey, apiSecret):
    # GetToken is the method used to authenticate to the FoD API, and extract the bearer token from the response
    authUrl = "https://api.ams.fortify.com/oauth/token"
    authorizationPayload = "scope=api-tenant&grant_type=client_credentials&client_id=" + apiKey + "&client_secret=" + apiSecret
    headers = {
        'content-type': "application/x-www-form-urlencoded",
        'cache-control': "no-cache"
    }
    response = requests.request("POST", authUrl, data=authorizationPayload, headers=headers)
    responseObject = json.loads(response.text)
    bearer = responseObject.get("access_token", "no token")

    if bearer != "no token":
        return responseObject['access_token']
    else:
        print(response.text)

    return None
def getDynamicConfig(releaseId, bearerToken):
    getDynamicConfigUrl = "https://api.ams.fortify.com/api/v3/releases/" + str(releaseId) + "/dynamic-scans/scan-setup"

    headers = {
        'authorization': "Bearer " + bearerToken,
        'Accept': "application/json"
    }
    response = requests.request("GET", getDynamicConfigUrl, headers=headers)
    try:
        if response.text != '{"success":false,"errors":["Application is not a Web / Thick-Client"]}':
            responseObject = json.loads(response.text)
        else:
            return None
    except:
        return None
    
    return responseObject
def generateReportRow(dynamicConfiguration, appName, rowCount):
    if(appName != ""):
        dynamicFacingType = dynamicConfiguration['dynamicScanEnvironmentFacingType']
        dynamicAuthType = dynamicConfiguration['dynamicScanAuthenticationType']
        primaryUserName = dynamicConfiguration['primaryUserName']
        secondaryUserName = dynamicConfiguration['secondaryUserName']
        otherUserName = dynamicConfiguration['otherUserName']
        siteUrl = dynamicConfiguration['dynamicSiteURL']
        timeZone = dynamicConfiguration['timeZone']
        entitlementFreqType = dynamicConfiguration['entitlementFrequencyType']

        worksheet.write(rowCount,0, appName)
        worksheet.write(rowCount,1, siteUrl)
        worksheet.write(rowCount,2, dynamicAuthType)    
        worksheet.write(rowCount,3, primaryUserName)
        worksheet.write(rowCount,4, secondaryUserName)
        worksheet.write(rowCount,5, otherUserName)    
        worksheet.write(rowCount,6, dynamicFacingType)
        worksheet.write(rowCount,7, entitlementFreqType)
        worksheet.write(rowCount,8, timeZone)

        print(appName + " Added successfully")  
      
    
getAllReleases(sys.argv[1], sys.argv[2])