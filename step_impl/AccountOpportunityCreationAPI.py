# from getgauge.python import step, Messages, data_store
# from simple_salesforce import Salesforce, SalesforceLogin
# from datetime import date
# import uuid
# import xlsxwriter
# from openpyxl import Workbook, load_workbook
# import csv
# import os
# import pytz
# from dateutil.parser import parse
# from dateutil import tz
# from datetime import timezone, tzinfo
# import datetime
# from random import randint
# import pdb
# from time import sleep
# # from jproperties import Properties, codecs
# from pathlib import Path
# import random
# import json

# # searchedData = sf.search("FIND {Automation}")
# # # print(searchedData['searchRecords'][1]['type'])
# # accountId = None
# # for records in searchedData['searchRecords']:
# #     if records['attributes']['type'] == 'Account':
# #         print(records)
# #         print(records['Id'])
# # #         accountId = records['attributes']['Id']
# #
# # print("Account ID: ", accountId)


# class AccountOpportunityCreation():    
#     session_id, instance = SalesforceLogin(
#         username=os.getenv("USER_ID"), password=os.getenv("USER_PASSWORD"), security_token=os.getenv("USER_SECURITY_TOKEN"), domain='test')
#     print(session_id, "\n", instance)
#     Messages.write_message(str(session_id) + " : " + str(instance))
#     sf = Salesforce(instance=instance, session_id=session_id)

#     def verifyATAB(self, createdByUserID, formattedISODate):
#         batchATAB = ["AccountTerritoryAssociationBatch", "UserHierarchyQueueDataSetupBatch", "UserHierarchyQueueDataSetupBatchUUHList", "UserHierarchyQueueDataSetupBatchO2AList", "UserHierarchyQueueAccountBatch", "UserHierarchyQueueOpportunityBatch",
#                      "UserHierarchyQueueGoalBatch", "SupportUserHierarchySharingBatch", "userSharingBatch", "UserSharingBatchForFormHeaders", "UserRevokeSharingDataSetUpBatch", "UserSharingRevokeBatch", "UserSharingRevokeForFormHeaders"]

#         batchATABStatus = {}
#         apexClassName = None
#         apexClassStatus = None
#         for batchValue in batchATAB:
#             print("Searching Current Batch:", batchValue)
#             isCompleted = False
#             while isCompleted == False:
#                 print("in while block")
#                 batchATABSOQL = f"SELECT CreatedDate, ApexClass.Name, MethodName, Status, CompletedDate FROM AsyncApexJob where ApexClass.Name = '{batchValue}' and CreatedById = '{createdByUserID}' and CreatedDate > {formattedISODate}"
#                 apexJobsResult = AccountOpportunityCreation.sf.query_all(
#                     query=batchATABSOQL)
#                 apexJobsData = apexJobsResult['records']
#                 print("\n", apexJobsData)
#                 if len(apexJobsData) > 0:
#                     if apexJobsData[0]['ApexClass'] != None:
#                         print(
#                             f"Verifying {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
#                         if apexJobsData[0]['Status'] == 'Completed':
#                             print(
#                                 f"Verified {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
#                             apexClassName = apexJobsData[0]['ApexClass']['Name']
#                             apexClassStatus = apexJobsData[0]['Status']
#                             batchATABStatus[apexClassName] = apexClassStatus
#                             isCompleted = True
#                             Messages.write_message(f"Verified {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
#                         elif apexJobsData[0]['Status'] == 'Failed':
#                             print(
#                                 f"Verified {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
#                             Messages.write_message(f"Verified {apexJobsData[0]['ApexClass']['Name']} -- {apexJobsData[0]['Status']}")
#                             apexClassName = apexJobsData[0]['ApexClass']['Name']
#                             apexClassStatus = apexJobsData[0]['Status']
#                             batchATABStatus[apexJobsData[0]
#                                             ['ApexClass']['Name']] = apexClassStatus
#                             isCompleted = True
#                         else:
#                             sleep(15)
#                 else:
#                     sleep(15)
#                     isCompleted = False
#         return batchATABStatus

#     def getDateTime(self):
#         currentDateTime = datetime.datetime.now(
#             pytz.timezone("America/Los_Angeles"))
#         twoMinutes = datetime.timedelta(minutes=2)
#         end = currentDateTime - twoMinutes

#         current_month = end.strftime('%m')
#         # current_month_text = end.strftime('%h')
#         # current_month_text = end.strftime('%B')

#         current_day = end.strftime('%d')
#         # current_day_text = end.strftime('%a')
#         # current_day_full_text = end.strftime('%A')

#         # current_weekday_day_of_today = end.strftime('%w')

#         current_year_full = end.strftime('%Y')
#         # current_year_short = end.strftime('%y')

#         current_second = end.strftime('%S')
#         current_minute = end.strftime('%M')
#         current_hour = end.strftime('%H')
#         # current_hour_word = end.strftime('%I')

#         # d = datetime.datetime.today().replace(microsecond=0)
#         # print(d.isoformat())

#         finalDateTime = datetime.datetime(int(current_year_full), int(current_month), int(current_day), int(
#             current_hour), int(current_minute), int(current_second), tzinfo=tz.gettz("America/Los_Angeles")).isoformat()
#         print(datetime.datetime(int(current_year_full), int(current_month), int(current_day), int(current_hour), int(
#             current_minute), int(current_second), tzinfo=tz.gettz("America/Los_Angeles")).isoformat())
        
#         Messages.write_message(f"ATAB RUN TIME: {finalDateTime}")
#         return finalDateTime

#     def getNAPIDate(self, sf, isPastOrFuture):
#         if isPastOrFuture == "Future":
#             soql = f"SELECT Is_Active__c,Id, NAPI_Date__c, Name FROM NAPI_Insert_date__c WHERE NAPI_Date__c > Today and Is_Active__c = true ORDER BY NAPI_Date__c ASC NULLS LAST"
#         elif isPastOrFuture == "Past":
#             soql = f"SELECT Id, NAPI_Date__c, Name FROM NAPI_Insert_date__c WHERE NAPI_Date__c < LAST_N_MONTHS:12 and Is_Active__c = true ORDER BY NAPI_Date__c DESC NULLS LAST"
#         queryResult = sf.query_all(query=soql)
#         recDetails = queryResult['records']
#     #     Messages.write_message("NAPI Date:" + str(recDetails[0]["Name"]))
#         print("NAPI Date: ", recDetails[0]["NAPI_Date__c"])
        
#         Messages.write_message(f"NAPI DATE: {recDetails[0]['NAPI_Date__c']}")
#         return str(recDetails[0]["Id"]), str(recDetails[0]["NAPI_Date__c"])

#     def getInstoreCycle(self, sf, cycleDate, current_month, current_year_full):
#         cycleSOQL = f"SELECT Id,Begin_Date__c,End_Date__c,Name,Week__c FROM Cycle__c where Name = '{cycleDate}' and Week__c = 1.0"

#         cycleResult = sf.query_all(query=cycleSOQL)
#         cycleData = cycleResult['records']

#         for cycle in cycleData:
#             if cycle['Name'] == f"{current_year_full}{current_month}" and cycle['Week__c'] == 1.0:
#                 print(cycle['Id'], "\t", cycle['Name'], "\t", cycle['Week__c'])        
#         Messages.write_message(f"INSTORE CYCLE: {cycle['Name']}")
#         return cycle['Id']

#     @step("Create <region> Account")
#     def create_account_and_store_to(self, region):
#         data_store.spec.clear()
#         # p = Properties()

#         # parentPath = Path(__file__).parents[1]
#         # accountPropertiesFileName = str(
#         #     parentPath) + "\\Data\\" + "Accounts.properties"

#         # with open(accountPropertiesFileName, "r+b") as f:
#         #     p.load(f, encoding='latin-1')

#         accountType = region
#         data_store.spec["REGION"] = region

#         # accountData = p.properties

#         if accountType == "USD":
#             currencyType = "USD"
#             assignedTerritories = os.getenv('US_TERRITORIES')
#             # assignedTerritories = 'LA Sales 2-6,Ssd Sales 2-3,NTL - COLGATE,Merchandising Region,Digital Division,C51 West Sales 4 - US'
#         elif accountType == "CAD":
#             currencyType = "CAD"
#             assignedTerritories = os.getenv('CA_TERRITORIES')
#             # assignedTerritories = 'Montreal Sales 2,Merchandising Canada Team,SSD Canada Sales 1,Digital Canada Sales 5,C51 Division - US'
#         assignedTerritoriesList = assignedTerritories.split(",")
#         Messages.write_message("ASSIGNED TERRITORIES: {assignedTerritoriesList}")

#         assignedTerritoriesMap = {}
#         for assignedTerritory in assignedTerritoriesList:
#             territoryNameSOQL = f"SELECT Id,Name,ParentTerritory2Id FROM Territory2 where Name = '{assignedTerritory}' and isActive__c = true"
#             territoryNameResult = AccountOpportunityCreation.sf.query_all(
#                 query=territoryNameSOQL)
#             territoryNameData = territoryNameResult['records']
#             if len(territoryNameData) > 0:
#                 territoryID = territoryNameData[0]['Id']
#                 assignedTerritoriesMap[assignedTerritory] = territoryNameData[0]['Id']
#         Messages.write_message(f"ASSIGNED TERRITORIES MAP: {assignedTerritoriesMap}")
#         accountName = f"AT-{str(randint(1, 99999))}"
#         isAccountCreated = AccountOpportunityCreation.sf.Account.create(
#             {'RecordTypeId': '012f1000000n6QJAAY', 'Name': f'{accountName}', 'CurrencyIsoCode': f'{currencyType}'})

#         if isAccountCreated["success"] == True:
#             print(f"Account {accountName} is created {isAccountCreated['id']}")
#             Messages.write_message(f"Account {accountName} is created {isAccountCreated['id']}")
#             data_store.spec[f'{region}_ACCOUNT_NAME'] = accountName
#             data_store.spec[f'{region}_ACCOUNT_ID'] = isAccountCreated['id']
#             # p['REGION'] = region
#             # p['ACCOUNT_NAME'] = accountName
#             # p['ACCOUNT_ID'] = isAccountCreated['id']

#             for assignedTerritory in assignedTerritoriesList:
#                 isObjectTerritory2Associated = AccountOpportunityCreation.sf.ObjectTerritory2Association.create(
#                     {'AssociationCause': 'Territory2Manual', 'ObjectId': isAccountCreated['id'], 'Territory2Id': f'{assignedTerritoriesMap[assignedTerritory]}'})
#                 print(isObjectTerritory2Associated)

#         # contactId = None
#         # contactData = sf.quick_search('fagoon4u@gmail.com')
#         # contactId = contactData['searchRecords'][0]['Id']
#         # print("Found Contact: ",contactData['searchRecords'][0]['Id'])

#         firstName = f"FN-{str(randint(1, 9999))}"
#         lastName = f"LN-{str(randint(1, 9999))}"
#         isContaactCreated = AccountOpportunityCreation.sf.Contact.create(
#             {'AccountId': isAccountCreated['id'], 'Email': 'fagoon4u@gmail.com', 'IsPrimaryContact__c': True, 'FirstName': firstName, 'LastName': lastName})
#         if isContaactCreated["success"] == True:
#             # p['CONTACT_NAME'] = f"{firstName} {lastName}"
#             # p['CONTACT_ID'] = isContaactCreated['id']
#             print(f"Contact {isContaactCreated['id']} is created...")
#             Messages.write_message(f"Contact {firstName} {lastName} is created {isContaactCreated['id']}")
#             data_store.spec[f'{region}_CONTACT_NAME'] = f"{firstName} {lastName}"
#             data_store.spec[f'{region}_CONTACT_ID'] = isContaactCreated['id']
#         # f.truncate(0)
#         # p.store(f, encoding='latin-1')
#         # pdb.set_trace()

#     @step("Run ATAB Batch")
#     def run_atab_batch(self):
#         userData = AccountOpportunityCreation.sf.quick_search(
#             f"{os.getenv('USER_ID')}")
#         userId = userData['searchRecords'][0]['Id']
#         print(userData['searchRecords'][0]['Id'])

#         result = AccountOpportunityCreation.sf.restful('tooling/executeAnonymous',
#                                                        {'anonymousBody': 'AccountTerritoryAssociationBatch a = new AccountTerritoryAssociationBatch(); Database.executeBatch(a, 25);'})
#         print(result)
#         Messages.write_message(f"{result}")
#         sleep(10)

#         formattedDateTime = self.getDateTime()
#         data_store.spec["ATAB_RUNNING_TIME"] = formattedDateTime
#         createdByUserID = userId
#         batchATABStatus = self.verifyATAB(createdByUserID, formattedDateTime)
#         print(batchATABStatus)

#     @step("Verify <region> Account Teams")
#     def verify_account_teams(self, region):
#         accountType = region
#         # p = Properties()
#         # parentPath = Path(__file__).parents[1]
#         # accountPropertiesFileName = str(
#         #     parentPath) + "\\Data\\" + "Accounts.properties"

#         # with open(accountPropertiesFileName, "r+b") as f:
#         #     p.load(f, encoding='latin-1')

#         #     accountData = p.properties
#         accountID = data_store.spec.get(f'{region}_ACCOUNT_ID')

#         if accountType == "USD":
#             assignedTerritories = os.getenv('US_TERRITORIES')
#         elif accountType == "CAD":
#             assignedTerritories = os.getenv('CA_TERRITORIES')
#         assignedTerritoriesList = assignedTerritories.split(",")

#         assignedTerritoriesMap = {}
#         for assignedTerritory in assignedTerritoriesList:
#             territoryNameSOQL = f"SELECT Id,Name,ParentTerritory2Id FROM Territory2 where Name = '{assignedTerritory}' and isActive__c = true"
#             territoryNameResult = AccountOpportunityCreation.sf.query_all(
#                 query=territoryNameSOQL)
#             territoryNameData = territoryNameResult['records']
#             if len(territoryNameData) > 0:
#                 territoryID = territoryNameData[0]['Id']
#                 # p[assignedTerritory] = territoryNameData[0]['Id']
#                 assignedTerritoriesMap[assignedTerritory] = territoryNameData[0]['Id']

#         usersVerified = []
#         usersNotVerified = []

#         for assignedTerritory in assignedTerritoriesList:
#             verifyMember = False
#             territoryMembersData = None
#             territoryIdToVerify = assignedTerritoriesMap[assignedTerritory]
#             print(f"----------------------------------------------------------------")
#             print(
#                 f"Territory to Verify: {territoryIdToVerify} : {assignedTerritory}")
#             print(f"----------------------------------------------------------------")
#             usersVerified = []
#             usersNotVerified = []
#             while verifyMember == False:
#                 territoryMembersSOQL = f"SELECT RoleInTerritory2,Territory2Id,UserId,User.Name FROM UserTerritory2Association where Territory2Id='{territoryIdToVerify}'"
#                 territoryMembersResult = AccountOpportunityCreation.sf.query_all(
#                     query=territoryMembersSOQL)
#                 territoryMembersData = territoryMembersResult['records']
#                 print(
#                     f"Total Members in Territory: {len(territoryMembersData)}")
#                 if len(territoryMembersData) > 0:
#                     for territoryMembers in territoryMembersData:
#                         if territoryMembers['RoleInTerritory2'] in ('Primary', 'Primary Split', 'Later Primary'):
#                             print(
#                                 f"{territoryMembers['User']['Name']} == {territoryMembers['RoleInTerritory2']}")
#                             Messages.write_message(f"{territoryMembers['User']['Name']} == {territoryMembers['RoleInTerritory2']}")
#                             verifyMember = True

#                 if len(territoryMembersData) == 0 or verifyMember == False:
#                     territoryNameSOQL = f"SELECT Id,Name,ParentTerritory2Id,ParentTerritory2.Name FROM Territory2 where Id = '{territoryIdToVerify}'"
#                     territoryNameResult = AccountOpportunityCreation.sf.query_all(
#                         query=territoryNameSOQL)
#                     territoryNameData = territoryNameResult['records']
#                     print(
#                         f"Record in Parent Territory: {len(territoryMembersData)}")
#                     if len(territoryNameData) > 0:
#                         territoryIdToVerify = territoryNameData[0]['ParentTerritory2Id']
#                         print(
#                             f"Parent Territory: {territoryNameData[0]['ParentTerritory2']['Name']}")
#                         Messages.write_message(f"Parent Territory: {territoryNameData[0]['ParentTerritory2']['Name']}")

#             print(f"Verify Member: {verifyMember}")

#             for territoryMembers in territoryMembersData:
#                 # '{accID}'"
#                 accountMembersSOQL = f"SELECT Role_In_Territory__c ,TeamMemberRole__c,TerritoryId__c,Territory_Category__c,User__c,User__r.Name FROM Account_Team__c where Account__c = '{accountID}'"
#                 accountMembersResult = AccountOpportunityCreation.sf.query_all(
#                     query=accountMembersSOQL)
#                 accountMembersData = accountMembersResult['records']
#         #         pdb.set_trace()
#                 isMemberVerified = False
#                 if len(accountMembersData) > 0:
#                     for accountMembers in accountMembersData:
#                         if not isMemberVerified:
#                             print(
#                                 "\nVerifying Users", territoryMembers['User']['Name'], "\t", accountMembers['User__r']['Name'])
#             #                 if (accountMembers['Role_In_Territory__c'] == territoryMembers['RoleInTerritory2']) and (accountMembers['TerritoryId__c'] == territoryMembers['Territory2Id']) and (accountMembers['User__c'] == territoryMembers['UserId']):
#                             if (accountMembers['Role_In_Territory__c'] in territoryMembers['RoleInTerritory2']) and (accountMembers['User__c'] == territoryMembers['UserId']):
#                                 print(
#                                     "\nMember Verified: ", accountMembers['User__r']['Name'], "\t", accountMembers['Role_In_Territory__c'])
#                                 Messages.write_message(f"Member Verified: {accountMembers['User__r']['Name']} \t {accountMembers['Role_In_Territory__c']}")
#                                 if territoryMembers['User']['Name'] not in usersVerified:
#                                     usersVerified.append(
#                                         territoryMembers['User']['Name'])

#                                 if territoryMembers['User']['Name'] in usersNotVerified:
#                                     usersNotVerified.remove(
#                                         territoryMembers['User']['Name'])
#                                 isMemberVerified = True
#                             else:
#                                 #                 print("Member Not Verified: ", accountMembers['User__r']['Name'], "\t", accountMembers['Role_In_Territory__c'])
#                                 if territoryMembers['User']['Name'] not in usersNotVerified:
#                                     usersNotVerified.append(
#                                         territoryMembers['User']['Name'])

#                                 if territoryMembers['User']['Name'] in usersVerified:
#                                     usersVerified.remove(
#                                         territoryMembers['User']['Name'])
#                                 isMemberVerified = False
#                 else:
#                     Messages.write_message("No Team Members Found...")
#                     print("No Team Members Found...")
#                 print("\nMembers Verified: ", usersVerified)
#                 print("\nMembers not verified: ", usersNotVerified)
#             # f.truncate(0)
#             # p.store(f, encoding='latin-1')

#     @step("Create Opportunities for <region> Account")
#     def create_opportunities_for_account(self, region):
#         end = datetime.datetime.now()
#         current_month = end.strftime('%m')
#         current_day = end.strftime('%d')
#         current_year_full = end.strftime('%Y')
#         # current_second = end.strftime('%S')
#         # current_minute = end.strftime('%M')
#         # current_hour = end.strftime('%H')

#         futureDate = datetime.datetime(int(current_year_full) + 1,
#                                        int(current_month), int(current_day)).date()
#         pastDate = datetime.datetime(int(current_year_full) - 1,
#                                      int(current_month), int(current_day)).date()
#         userData = AccountOpportunityCreation.sf.quick_search(
#             f"{os.getenv('USER_ID')}")
#         userId = userData['searchRecords'][0]['Id']
#         print(userData['searchRecords'][0]['Id'])
#         data_store.spec["FUTURE_DATE"] = futureDate
#         data_store.spec["PAST_DATE"] = pastDate
#         data_store.spec["USER_ID"] = userId

#         napiDateId, napiDate = AccountOpportunityCreation.getNAPIDate(
#             self, AccountOpportunityCreation.sf, "Future")
#         data_store.spec["FUTURE_NAPI_DATE"] = napiDate
        
#         pastNapiDateId, pastNAPIDate = AccountOpportunityCreation.getNAPIDate(
#             self, AccountOpportunityCreation.sf, "Past")
#         data_store.spec["PAST_NAPI_DATE"] = pastNAPIDate
        
#         instoreCycle = AccountOpportunityCreation.getInstoreCycle(self,
#                                                        AccountOpportunityCreation.sf, f"{str(int(current_year_full)+1)}{current_month}", current_month, current_year_full)
#         pastInStoreCycle = AccountOpportunityCreation.getInstoreCycle(self,
#                                                                       AccountOpportunityCreation.sf, f"{str(int(current_year_full)-1)}{current_month}", current_month, current_year_full)
#         data_store.spec["FUTURE_INSTORE_CYCLE"] = instoreCycle
#         data_store.spec["PAST_INSTORE_CYCLE"] = pastInStoreCycle
        
#         productLineName = os.getenv('OPPORTUNITY_FOR_PRODUCT_LINE')
#         data_store.spec["PRODUCT_LINES"] = productLineName

#         usProductLineMap = {}
#         caProductLineMap = {}
#         productLineList = []
#         childOpps = []

#         accountID = data_store.spec.get(f'{region}_ACCOUNT_ID')

#         if region == 'USD':
#             currencyCode = 'USD'
#             if productLineName == "ALL":
#                 productLine = os.getenv('US_PRODUCT_LINES')
#             else:
#                 productLine = productLineName

#         elif region == 'CAD':
#             currencyCode = 'CAD'
#             if productLineName == "ALL":
#                 productLine = os.getenv('CA_PRODUCT_LINES')
#             else:
#                 productLine = productLineName

#         if "," in productLine:
#             productLineList = productLine.split(",")
#         else:
#             productLineList = [productLine]

#         # Get all opportunity record types :
#         opprtunityPLMap = {}
#         opprtunityPLSOQL = f"select Id,Name from RecordType where sObjectType='Opportunity'"
#         opprtunityPLResult = AccountOpportunityCreation.sf.query_all(
#             query=opprtunityPLSOQL)
#         opprtunityPLData = opprtunityPLResult['records']
#         for opprtunityPL in opprtunityPLData:
#             opprtunityPLMap[opprtunityPL['Name']] = opprtunityPL['Id']
#         print(opprtunityPLMap)
#         [print(key, value) for key, value in opprtunityPLMap.items()]
#         for productLine in productLineList:
#             print(
#                 f"\n{productLine} Started---------------------------------------------------------------------------\n")
#             usProductLineMap[productLine] = {
#                 'PRODUCT_LINE_ID': opprtunityPLMap[productLine], 'PRODUCT_LINE_NAME': productLine, 'PARENT_OPP': ''}

#             parentOpportunityName = f"{productLine}#P#{str(randint(1, 99999))}"
#             if productLine in ('InStore', 'InStore-Canada'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'AccountId': accountID})
#             elif productLine in ('FSI'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'AccountId': accountID})
#         #     elif productLine in ('Digital','Digital- Canada'):
#         #         parentOppDetails = sf.Opportunity.create({'RecordTypeId' : opprtunityPLMap[productLine],'Opportunity_Category__c' : 'Parent', 'Name' : parentOpportunityName,'StageName' : 'Contract','Insert_Date__c' : str(futureDate),'CloseDate' : str(futureDate), 'End_Date__c' : str(futureDate), 'AccountId': accountID})
#             else:
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID})

#             print(parentOppDetails)
#             if parentOppDetails["success"] == True:
#                 print(f"Parent Opp ID of {productLine}: ",
#                       parentOppDetails['id'])
#                 Messages.write_message(f"Parent Opp ID of {productLine}: {parentOppDetails['id']}")
#                 data_store.spec[f"1_PO_ID_{productLine}"] = parentOppDetails['id']
#                 data_store.spec[f"1_PO_NAME_{productLine}"] = parentOpportunityName
#                 usProductLineMap[productLine]['PARENT_OPP_1'] = {
#                     'ID': parentOppDetails['id'], 'CHILD1_OPP_ID': '', 'CHILD2_OPP_ID': '', 'CHILD3_OPP_ID': ''}
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD1_OPP_ID'] = {'ID':'','ORDER#' : '','PARENT_ORDER#' : ''}
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD2_OPP_ID'] = {'ID':'','ORDER#' : '','PARENT_ORDER#' : ''}
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD3_OPP_ID'] = {'ID':'','ORDER#' : '','PARENT_ORDER#' : ''}          
#             child1OpportunityName = f"{productLine}#FD#{str(randint(1, 99999))}"
#             p1OrderNumber = str(uuid.uuid4()).upper()[:10]
#             if productLine in ('InStore', 'InStore-Canada'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'Artwork_Due_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('FSI'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'End_Date__c': str(
#                     napiDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('SSMG'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Checkout 51'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Merchandising'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             else:
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             print(child1OppDetails)
#             if child1OppDetails["success"] == True:
#                 print(f"Child1 Opp ID of {productLine}: ",
#                       child1OppDetails['id'])
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD1_OPP_ID']['ID'] = child1OppDetails['id']
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD1_OPP_ID']['ORDER#'] = str(p1OrderNumber)
#                 childOpps.append(child1OppDetails['id'])
#                 Messages.write_message(f"Child 1 Opp ID of {productLine}: {child1OppDetails['id']}")
#                 data_store.spec[f"P1_C1_ID_{productLine}"] = child1OppDetails['id']
#                 data_store.spec[f"P1_C1_NAME_{productLine}"] = child1OpportunityName

#             child2OpportunityName = f"{productLine}#BD#{str(randint(1, 99999))}"
#             if productLine in ('InStore', 'InStore-Canada'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(pastInStoreCycle), 'CloseDate': str(
#                     pastDate), 'Artwork_Due_Date__c': str(pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('FSI'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(pastNapiDateId), 'CloseDate': str(pastNAPIDate), 'End_Date__c': str(
#                     pastNAPIDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('SSMG'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(pastDate), 'End_Date__c': str(
#                     pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Checkout 51'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Merchandising'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             else:
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             print(child2OppDetails)
#             if child2OppDetails["success"] == True:
#                 print(f"Child2 Opp ID of {productLine}: ",
#                       child2OppDetails['id'])
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD2_OPP_ID']['ID'] = child2OppDetails['id']
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD2_OPP_ID']['ORDER#'] = str(p1OrderNumber)
#                 childOpps.append(child2OppDetails['id'])
#                 Messages.write_message(f"Child 2 Opp ID of {productLine}: {child2OppDetails['id']}")
#                 data_store.spec[f"P1_C2_ID_{productLine}"] = child2OppDetails['id']
#                 data_store.spec[f"P1_C2_NAME_{productLine}"] = child2OpportunityName

#             child3OpportunityName = f"{productLine}#BD#{str(randint(1, 99999))}"
#             if productLine in ('InStore', 'InStore-Canada'):
#                 child3OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child3OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(pastInStoreCycle), 'CloseDate': str(
#                     pastDate), 'Artwork_Due_Date__c': str(pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('FSI'):
#                 child3OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child3OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(pastNapiDateId), 'CloseDate': str(pastNAPIDate), 'End_Date__c': str(
#                     pastNAPIDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('SSMG'):
#                 child3OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child3OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(pastDate), 'End_Date__c': str(
#                     pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Checkout 51'):
#                 child3OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child3OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Merchandising'):
#                 child3OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child3OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             else:
#                 child3OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child3OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : p1OrderNumber, 'Parent_Order__c' : ''})
#             print(child3OppDetails)
#             if child2OppDetails["success"] == True:
#                 print(f"Child3 Opp ID of {productLine}: ",
#                       child3OppDetails['id'])
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD3_OPP_ID']['ID'] = child3OppDetails['id']
#                 usProductLineMap[productLine]['PARENT_OPP_1']['CHILD3_OPP_ID']['ORDER#'] = str(p1OrderNumber)
#                 childOpps.append(child3OppDetails['id'])
#                 Messages.write_message(f"Child 3 Opp ID of {productLine}: {child3OppDetails['id']}")
#                 data_store.spec[f"P1_C3_ID_{productLine}"] = child3OppDetails['id']
#                 data_store.spec[f"P1_C3_NAME_{productLine}"] = child3OpportunityName                
# # --------------------------------------------------------------------------------------------------------------------------------------------------------------------

#             parentOpportunityName = f"{productLine}#P#{str(randint(1, 99999))}"
#             if productLine in ('InStore', 'InStore-Canada'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'AccountId': accountID})
#             elif productLine in ('FSI'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'AccountId': accountID})
#         #     elif productLine in ('Digital','Digital- Canada'):
#         #         parentOppDetails = sf.Opportunity.create({'RecordTypeId' : opprtunityPLMap[productLine],'Opportunity_Category__c' : 'Parent', 'Name' : parentOpportunityName,'StageName' : 'Contract','Insert_Date__c' : str(futureDate),'CloseDate' : str(futureDate), 'End_Date__c' : str(futureDate), 'AccountId': accountID})
#             else:
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID})

#             print(parentOppDetails)
#             if parentOppDetails["success"] == True:
#                 print(f"Parent Opp ID of {productLine}: ",
#                       parentOppDetails['id'])
#                 Messages.write_message(f"Parent Opp ID of {productLine}: {parentOppDetails['id']}")
#                 data_store.spec[f"2_PO_ID_{productLine}"] = parentOppDetails['id']
#                 data_store.spec[f"2_PO_NAME_{productLine}"] = parentOpportunityName
#                 usProductLineMap[productLine]['PARENT_OPP_2'] = {
#                     'ID': parentOppDetails['id'], 'CHILD1_OPP_ID': '', 'CHILD2_OPP_ID': ''}
#                 usProductLineMap[productLine]['PARENT_OPP_2']['CHILD1_OPP_ID'] = {'ID':'','ORDER#' : '','PARENT_ORDER#' : ''}
#                 usProductLineMap[productLine]['PARENT_OPP_2']['CHILD2_OPP_ID'] = {'ID':'','ORDER#' : '','PARENT_ORDER#' : ''}              
          
#             child1OpportunityName = f"{productLine}#FD#{str(randint(1, 99999))}"
#             child1OpportunityOrderNumber = str(uuid.uuid4()).upper()[:10]
#             if productLine in ('InStore', 'InStore-Canada'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'Artwork_Due_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : child1OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             elif productLine in ('FSI'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'End_Date__c': str(
#                     napiDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : child1OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             elif productLine in ('SSMG'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : child1OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             elif productLine in ('Checkout 51'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : child1OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             elif productLine in ('Merchandising'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : child1OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             else:
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : child1OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             print(child1OppDetails)
#             if child1OppDetails["success"] == True:
#                 print(f"Child1 Opp ID of {productLine}: ",
#                       child1OppDetails['id'])
#                 usProductLineMap[productLine]['PARENT_OPP_2']['CHILD1_OPP_ID']['ID'] = child1OppDetails['id']
#                 usProductLineMap[productLine]['PARENT_OPP_2']['CHILD1_OPP_ID']['ORDER#'] = str(child1OpportunityOrderNumber)
#                 usProductLineMap[productLine]['PARENT_OPP_2']['CHILD1_OPP_ID']['PARENT_ORDER#'] = str(p1OrderNumber)
#                 childOpps.append(child1OppDetails['id'])
#                 Messages.write_message(f"Child 1 Opp ID of {productLine}: {child1OppDetails['id']}")
#                 data_store.spec[f"P2_C1_ID_{productLine}"] = child1OppDetails['id']
#                 data_store.spec[f"P2_C1_NAME_{productLine}"] = child1OpportunityName

#             child2OpportunityName = f"{productLine}#BD#{str(randint(1, 99999))}"
#             child2OpportunityOrderNumber = str(uuid.uuid4()).upper()[:10]
#             if productLine in ('InStore', 'InStore-Canada'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(pastInStoreCycle), 'CloseDate': str(
#                     pastDate), 'Artwork_Due_Date__c': str(pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             elif productLine in ('FSI'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(pastNapiDateId), 'CloseDate': str(pastNAPIDate), 'End_Date__c': str(
#                     pastNAPIDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             elif productLine in ('SSMG'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(pastDate), 'End_Date__c': str(
#                     pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             elif productLine in ('Checkout 51'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             elif productLine in ('Merchandising'):
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             else:
#                 child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#                     pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1OrderNumber})
#             print(child2OppDetails)
#             if child2OppDetails["success"] == True:
#                 print(f"Child2 Opp ID of {productLine}: ",
#                       child2OppDetails['id'])
#                 usProductLineMap[productLine]['PARENT_OPP_2']['CHILD2_OPP_ID']['ID'] = child2OppDetails['id']
#                 usProductLineMap[productLine]['PARENT_OPP_2']['CHILD2_OPP_ID']['ORDER#'] = str(child2OpportunityOrderNumber)
#                 usProductLineMap[productLine]['PARENT_OPP_2']['CHILD2_OPP_ID']['PARENT_ORDER#'] = str(p1OrderNumber)     
#                 childOpps.append(child2OppDetails['id'])
#                 Messages.write_message(f"Child 2 Opp ID of {productLine}: {child2OppDetails['id']}")
#                 data_store.spec[f"P2_C2_ID_{productLine}"] = child2OppDetails['id']
#                 data_store.spec[f"P2_C2_NAME_{productLine}"] = child2OpportunityName


# # --------------------------------------------------------------------------------------------------------------------------------------------------------------------

#             parentOpportunityName = f"{productLine}#P#{str(randint(1, 99999))}"
#             if productLine in ('InStore', 'InStore-Canada'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'AccountId': accountID})
#             elif productLine in ('FSI'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'AccountId': accountID})
#         #     elif productLine in ('Digital','Digital- Canada'):
#         #         parentOppDetails = sf.Opportunity.create({'RecordTypeId' : opprtunityPLMap[productLine],'Opportunity_Category__c' : 'Parent', 'Name' : parentOpportunityName,'StageName' : 'Contract','Insert_Date__c' : str(futureDate),'CloseDate' : str(futureDate), 'End_Date__c' : str(futureDate), 'AccountId': accountID})
#             else:
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID})

#             print(parentOppDetails)
#             if parentOppDetails["success"] == True:
#                 print(f"Parent Opp ID of {productLine}: ",
#                       parentOppDetails['id'])
#                 Messages.write_message(f"Parent Opp ID of {productLine}: {parentOppDetails['id']}")
#                 data_store.spec[f"3_PO_ID_{productLine}"] = parentOppDetails['id']
#                 data_store.spec[f"3_PO_NAME_{productLine}"] = parentOpportunityName
#                 usProductLineMap[productLine]['PARENT_OPP_3'] = {
#                     'ID': parentOppDetails['id'], 'CHILD1_OPP_ID': ''}
#                 usProductLineMap[productLine]['PARENT_OPP_3']['CHILD1_OPP_ID'] = {'ID':'','ORDER#' : '','PARENT_ORDER#' : ''}
          
#             child1OpportunityName = f"{productLine}#FD#{str(randint(1, 99999))}"
#             p3Child1OpportunityOrderNumber = str(uuid.uuid4()).upper()[:10]
#             if productLine in ('InStore', 'InStore-Canada'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'Artwork_Due_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : p3Child1OpportunityOrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('FSI'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'End_Date__c': str(
#                     napiDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : p3Child1OpportunityOrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('SSMG'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : p3Child1OpportunityOrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Checkout 51'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : p3Child1OpportunityOrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Merchandising'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : p3Child1OpportunityOrderNumber, 'Parent_Order__c' : ''})
#             else:
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : p3Child1OpportunityOrderNumber, 'Parent_Order__c' : ''})
#             print(child1OppDetails)
#             if child1OppDetails["success"] == True:
#                 print(f"Child1 Opp ID of {productLine}: ",
#                       child1OppDetails['id'])
#                 usProductLineMap[productLine]['PARENT_OPP_3']['CHILD1_OPP_ID']['ID'] = child1OppDetails['id']
#                 usProductLineMap[productLine]['PARENT_OPP_3']['CHILD1_OPP_ID']['ORDER#'] = str(p3Child1OpportunityOrderNumber)
#                 childOpps.append(child1OppDetails['id'])
#                 Messages.write_message(f"Child 1 Opp ID of {productLine}: {child1OppDetails['id']}")
#                 data_store.spec[f"P3_C1_ID_{productLine}"] = child1OppDetails['id']
#                 data_store.spec[f"P3_C1_NAME_{productLine}"] = child1OpportunityName

#             # child2OpportunityName = f"{productLine}#BD#{str(randint(1, 99999))}"
#             # child2OpportunityOrderNumber = str(uuid.uuid4()).upper()[:10]
#             # if productLine in ('InStore', 'InStore-Canada'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(pastInStoreCycle), 'CloseDate': str(
#             #         pastDate), 'Artwork_Due_Date__c': str(pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1Child2OpportunityOrderNumber})
#             # elif productLine in ('FSI'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(pastNapiDateId), 'CloseDate': str(pastNAPIDate), 'End_Date__c': str(
#             #         pastNAPIDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1Child2OpportunityOrderNumber})
#             # elif productLine in ('SSMG'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(pastDate), 'End_Date__c': str(
#             #         pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1Child2OpportunityOrderNumber})
#             # elif productLine in ('Checkout 51'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#             #         pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1Child2OpportunityOrderNumber})
#             # elif productLine in ('Merchandising'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#             #         pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1Child2OpportunityOrderNumber})
#             # else:
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#             #         pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : child2OpportunityOrderNumber, 'Parent_Order__c' : p1Child2OpportunityOrderNumber})
#             # print(child2OppDetails)
#             # if child2OppDetails["success"] == True:
#             #     print(f"Child2 Opp ID of {productLine}: ",
#             #           child2OppDetails['id'])
#             #     usProductLineMap[productLine]['PARENT_OPP_3']['CHILD2_OPP_ID'] = child2OppDetails['id']
#             #     childOpps.append(child2OppDetails['id'])
#             #     Messages.write_message(f"Child 2 Opp ID of {productLine}: {child2OppDetails['id']}")
#             #     data_store.spec[f"P_OPP#3+C_OPP#2+ID+{productLine}"] = child2OppDetails['id']
#             #     data_store.spec[f"P_OPP#3+C_OPP#2+ID+{productLine}"] = child2OpportunityName

# # --------------------------------------------------------------------------------------------------------------------------------------------------------------------

#             parentOpportunityName = f"{productLine}#P#{str(randint(1, 99999))}"
#             if productLine in ('InStore', 'InStore-Canada'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'AccountId': accountID})
#             elif productLine in ('FSI'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'AccountId': accountID})
#         #     elif productLine in ('Digital','Digital- Canada'):
#         #         parentOppDetails = sf.Opportunity.create({'RecordTypeId' : opprtunityPLMap[productLine],'Opportunity_Category__c' : 'Parent', 'Name' : parentOpportunityName,'StageName' : 'Contract','Insert_Date__c' : str(futureDate),'CloseDate' : str(futureDate), 'End_Date__c' : str(futureDate), 'AccountId': accountID})
#             else:
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID})

#             print(parentOppDetails)
#             if parentOppDetails["success"] == True:
#                 print(f"Parent Opp ID of {productLine}: ",
#                       parentOppDetails['id'])
#                 Messages.write_message(f"Parent Opp ID of {productLine}: {parentOppDetails['id']}")
#                 data_store.spec[f"4_PO_ID_{productLine}"] = parentOppDetails['id']
#                 data_store.spec[f"4_PO_NAME_{productLine}"] = parentOpportunityName
#                 usProductLineMap[productLine]['PARENT_OPP_4'] = {
#                     'ID': parentOppDetails['id'], 'CHILD1_OPP_ID': ''}
#                 usProductLineMap[productLine]['PARENT_OPP_4']['CHILD1_OPP_ID'] = {'ID':'','ORDER#' : '','PARENT_ORDER#' : ''}
          
#             child1OpportunityName = f"{productLine}#FD#CO#{str(randint(1, 99999))}"
#             orderNumber = str(uuid.uuid4()).upper()[:10]
#             if productLine in ('InStore', 'InStore-Canada'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'Artwork_Due_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : orderNumber, 'Parent_Order__c' : p3Child1OpportunityOrderNumber})
#             elif productLine in ('FSI'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'End_Date__c': str(
#                     napiDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : orderNumber, 'Parent_Order__c' : p3Child1OpportunityOrderNumber})
#             elif productLine in ('SSMG'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : orderNumber, 'Parent_Order__c' : p3Child1OpportunityOrderNumber})
#             elif productLine in ('Checkout 51'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : orderNumber, 'Parent_Order__c' : p3Child1OpportunityOrderNumber})
#             elif productLine in ('Merchandising'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : orderNumber, 'Parent_Order__c' : p3Child1OpportunityOrderNumber})
#             else:
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : orderNumber, 'Parent_Order__c' : p3Child1OpportunityOrderNumber})
#             print(child1OppDetails)
#             if child1OppDetails["success"] == True:
#                 print(f"Child1 Opp ID of {productLine}: ",
#                       child1OppDetails['id'])
#                 usProductLineMap[productLine]['PARENT_OPP_4']['CHILD1_OPP_ID']['ID'] = child1OppDetails['id']
#                 usProductLineMap[productLine]['PARENT_OPP_4']['CHILD1_OPP_ID']['ORDER#'] = str(orderNumber)
#                 usProductLineMap[productLine]['PARENT_OPP_4']['CHILD1_OPP_ID']['PARENT_ORDER#'] = str(p3Child1OpportunityOrderNumber)
#                 childOpps.append(child1OppDetails['id'])
#                 Messages.write_message(f"Child 1 Opp ID of {productLine}: {child1OppDetails['id']}")
#                 data_store.spec[f"P4_C1_ID_{productLine}"] = child1OppDetails['id']
#                 data_store.spec[f"P4_C1_NAME_{productLine}"] = child1OpportunityName

#             # child2OpportunityName = f"{productLine}#BD#{str(randint(1, 99999))}"
#             # # child2OpportunityOrderNumber = str(uuid.uuid4()).upper()[:10]
#             # if productLine in ('InStore', 'InStore-Canada'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(pastInStoreCycle), 'CloseDate': str(
#             #         pastDate), 'Artwork_Due_Date__c': str(pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : orderNumber, 'Parent_Order__c' : ''})
#             # elif productLine in ('FSI'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(pastNapiDateId), 'CloseDate': str(pastNAPIDate), 'End_Date__c': str(
#             #         pastNAPIDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : orderNumber, 'Parent_Order__c' : ''})
#             # elif productLine in ('SSMG'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(pastDate), 'End_Date__c': str(
#             #         pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : orderNumber, 'Parent_Order__c' : ''})
#             # elif productLine in ('Checkout 51'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#             #         pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : orderNumber, 'Parent_Order__c' : ''})
#             # elif productLine in ('Merchandising'):
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#             #         pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : orderNumber, 'Parent_Order__c' : ''})
#             # else:
#             #     child2OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child2OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(pastDate), 'CloseDate': str(
#             #         pastDate), 'End_Date__c': str(pastDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : orderNumber, 'Parent_Order__c' : ''})
#             # print(child2OppDetails)
#             # if child2OppDetails["success"] == True:
#             #     print(f"Child2 Opp ID of {productLine}: ",
#             #           child2OppDetails['id'])
#             #     usProductLineMap[productLine]['PARENT_OPP_4']['CHILD2_OPP_ID'] = child2OppDetails['id']
#             #     childOpps.append(child2OppDetails['id'])
#             #     Messages.write_message(f"Child 2 Opp ID of {productLine}: {child2OppDetails['id']}")
#             #     data_store.spec[f"P_OPP#4+C_OPP#2+ID+{productLine}"] = child2OppDetails['id']
#             #     data_store.spec[f"P_OPP#4+C_OPP#2+ID+{productLine}"] = child2OpportunityName

# # --------------------------------------------------------------------------------------------------------------------------------------------------------------------

#             parentOpportunityName = f"{productLine}#P#{str(randint(1, 99999))}"
#             if productLine in ('InStore', 'InStore-Canada'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'AccountId': accountID})
#             elif productLine in ('FSI'):
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'AccountId': accountID})
#         #     elif productLine in ('Digital','Digital- Canada'):
#         #         parentOppDetails = sf.Opportunity.create({'RecordTypeId' : opprtunityPLMap[productLine],'Opportunity_Category__c' : 'Parent', 'Name' : parentOpportunityName,'StageName' : 'Contract','Insert_Date__c' : str(futureDate),'CloseDate' : str(futureDate), 'End_Date__c' : str(futureDate), 'AccountId': accountID})
#             else:
#                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
#                                                                                      'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID})

#             print(parentOppDetails)
#             if parentOppDetails["success"] == True:
#                 print(f"Parent Opp ID of {productLine}: ",
#                       parentOppDetails['id'])
#                 Messages.write_message(f"Parent Opp ID of {productLine}: {parentOppDetails['id']}")
#                 data_store.spec[f"5_PO_ID_{productLine}"] = parentOppDetails['id']
#                 data_store.spec[f"5_PO_NAME_{productLine}"] = parentOpportunityName
#                 usProductLineMap[productLine]['PARENT_OPP_5'] = {
#                     'ID': parentOppDetails['id'], 'CHILD1_OPP_ID': ''}
#                 usProductLineMap[productLine]['PARENT_OPP_5']['CHILD1_OPP_ID'] = {'ORDER#' : '','PARENT_ORDER#' : ''}
          
#             child1OpportunityName = f"{productLine}#FD#NPO#{str(randint(1, 99999))}"
#             childOrderNumber = str(uuid.uuid4()).upper()[:10]
#             if productLine in ('InStore', 'InStore-Canada'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'Artwork_Due_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('FSI'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'End_Date__c': str(
#                     napiDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('SSMG'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(
#                     futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Checkout 51'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
#             elif productLine in ('Merchandising'):
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
#             else:
#                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
#                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
#             print(child1OppDetails)
#             if child1OppDetails["success"] == True:
#                 print(f"Child1 Opp ID of {productLine}: ",
#                       child1OppDetails['id'])
#                 usProductLineMap[productLine]['PARENT_OPP_5']['CHILD1_OPP_ID']['ID'] = child1OppDetails['id']
#                 usProductLineMap[productLine]['PARENT_OPP_5']['CHILD1_OPP_ID']['ORDER#'] = str(childOrderNumber)
#                 childOpps.append(child1OppDetails['id'])
#                 Messages.write_message(f"Child 1 Opp ID of {productLine}: {child1OppDetails['id']}")
#                 data_store.spec[f"P5_C1_ID_{productLine}"] = child1OppDetails['id']
#                 data_store.spec[f"P5_C1_NAME_{productLine}"] = child1OpportunityName

# # # --------------------------------------------------------------------------------------------------------------------------------------------------------------------

# #             parentOpportunityName = f"{productLine}#P#{str(randint(1, 99999))}"
# #             if productLine in ('InStore', 'InStore-Canada'):
# #                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
# #                                                                                      'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'AccountId': accountID})
# #             elif productLine in ('FSI'):
# #                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
# #                                                                                      'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'AccountId': accountID})
# #         #     elif productLine in ('Digital','Digital- Canada'):
# #         #         parentOppDetails = sf.Opportunity.create({'RecordTypeId' : opprtunityPLMap[productLine],'Opportunity_Category__c' : 'Parent', 'Name' : parentOpportunityName,'StageName' : 'Contract','Insert_Date__c' : str(futureDate),'CloseDate' : str(futureDate), 'End_Date__c' : str(futureDate), 'AccountId': accountID})
# #             else:
# #                 parentOppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Parent', 'Name': parentOpportunityName,
# #                                                                                      'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID})

# #             print(parentOppDetails)
# #             if parentOppDetails["success"] == True:
# #                 print(f"Parent Opp ID of {productLine}: ",
# #                       parentOppDetails['id'])
# #                 Messages.write_message(f"Parent Opp ID of {productLine}: {parentOppDetails['id']}")
# #                 data_store.spec[f"6_PO_ID_{productLine}"] = parentOppDetails['id']
# #                 data_store.spec[f"6_PO_NAME_{productLine}"] = parentOpportunityName
# #                 usProductLineMap[productLine]['PARENT_OPP_6'] = {
# #                     'ID': parentOppDetails['id'], 'CHILD1_OPP_ID': ''}
          
# #             child1OpportunityName = f"{productLine}#FD#NPO#{str(randint(1, 99999))}"
# #             childOrderNumber = str(uuid.uuid4()).upper()[:10]
# #             if productLine in ('InStore', 'InStore-Canada'):
# #                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'InStore_Cycle__c': str(instoreCycle), 'CloseDate': str(futureDate), 'End_Date__c': str(
# #                     futureDate), 'Artwork_Due_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPS__c': 1, 'Estimated_Store_Count__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Tactical', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
# #             elif productLine in ('FSI'):
# #                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'NAPI_Insert_date__c': str(napiDateId), 'CloseDate': str(napiDate), 'End_Date__c': str(
# #                     napiDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'ILocCirculationCharges__c': '[{"ChargeType":"C","Charges":1000,"Description":"CIRCULATION CHARGE\'s","Amount":1,"CirculationQty":1000}]', 'ILOCProductionCharges__c': '[{"ChargeType":"P","Charges":2000,"Description":"DISK HANDLING\'s","Amount":1000,"CirculationQty":2}]', 'ILocOtherCharges__c': '[{"ChargeType":"o","Charges":5000,"Description":"OTHER CHARGE\'s","Amount":50, "CirculationQty": 100}]', 'ILOCTotalProgramFee__c': 8000, 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
# #             elif productLine in ('SSMG'):
# #                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(futureDate), 'End_Date__c': str(
# #                     futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Estimated_Average_CPM__c': 1, 'Expected_Circulation__c': 1, 'Business_Type__c': 'New', 'Type': 'Standard', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
# #             elif productLine in ('Checkout 51'):
# #                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
# #                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
# #             elif productLine in ('Merchandising'):
# #                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
# #                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Subscription', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
# #             else:
# #                 child1OppDetails = AccountOpportunityCreation.sf.Opportunity.create({'RecordTypeId': opprtunityPLMap[productLine], 'Opportunity_Category__c': 'Child', 'Name': child1OpportunityName, 'StageName': 'Contract', 'Insert_Date__c': str(futureDate), 'CloseDate': str(
# #                     futureDate), 'End_Date__c': str(futureDate), 'AccountId': accountID, 'Parent_Opportunity__c': parentOppDetails['id'], 'Probability__c': '75', 'Business_Type__c': 'New', 'Status__c': 'Reserved-RS1', 'Type': 'Standard', 'Order__c' : childOrderNumber, 'Parent_Order__c' : ''})
# #             print(child1OppDetails)
# #             if child1OppDetails["success"] == True:
# #                 print(f"Child1 Opp ID of {productLine}: ",
# #                       child1OppDetails['id'])
# #                 usProductLineMap[productLine]['PARENT_OPP_6']['CHILD1_OPP_ID'] = child1OppDetails['id']
# #                 childOpps.append(child1OppDetails['id'])
# #                 Messages.write_message(f"Child 1 Opp ID of {productLine}: {child1OppDetails['id']}")
# #                 data_store.spec[f"10_C_ID_{productLine}"] = child1OppDetails['id']
# #                 data_store.spec[f"10_C_NAME_{productLine}"] = child1OpportunityName
                                                               
#             print("\nChild Opportunities\n", childOpps)

# #-------------------------------------------------------------------------------------------------------------------------------------
#             productsList = []
#             productsList.clear()
#             if productLine in ('InStore', 'InStore-Canada'):
#                 oppPLProductsSOQL = f"SELECT Id,Product_Line__c, Product_Line__r.Name,Product__c, Product__r.Name FROM Product_Junction__c where Product_Line__r.Name = 'InStore' and Product__r.IsActive = true and Product__c in (SELECT Product2Id from PricebookEntry where CurrencyIsoCode = '{currencyCode}' and IsActive = true)"
#             else:
#                 oppPLProductsSOQL = f"SELECT Id,Product_Line__c, Product_Line__r.Name,Product__c, Product__r.Name FROM Product_Junction__c where Product_Line__r.Name = '{productLine}' and Product__r.IsActive = true and Product__c in (SELECT Product2Id from PricebookEntry where CurrencyIsoCode = '{currencyCode}' and IsActive = true)"
#         #         oppPLProductsSOQL = f"SELECT Id,Product_Line__c, Product_Line__r.Name,Product__c, Product__r.Name FROM Product_Junction__c where Product_Line__r.Name = '{productLine}'"
#             oppPLProductsResult = AccountOpportunityCreation.sf.query_all(
#                 query=oppPLProductsSOQL)
#             oppPLProductsData = oppPLProductsResult['records']
            
#             for oppPLProduct in oppPLProductsData:
#                 if oppPLProduct['Product__r'] != None:
#                     if 'Name' in oppPLProduct['Product__r']:
#                         productsList.append(oppPLProduct['Product__r']['Name'])
#                         print("Found Product: ",
#                               oppPLProduct['Product__r']['Name'])
#         #     pdb.set_trace()
#             # print("\nProducts\n", productsList)
#             print(*productsList, sep = "\n")
#             IsActive = False
#             while IsActive == False:
#                 productName = random.choice(productsList)
                
#                 Messages.write_message(f"Product Name: {productName}")
#                 data_store.spec[f"PRODUCT_NAME"] = productName
                
#                 print("Product Name: ", productName)
#                 # priceBookSOQL = f"SELECT Id, Name,IsActive FROM PricebookEntry where Name = '{productName}' and CurrencyIsoCode = '{currencyCode}' and IsActive = true"
#                 priceBookSOQL = f"SELECT Id, IsActive FROM PricebookEntry WHERE CurrencyIsoCode = '{currencyCode}' AND IsActive = True AND Product2.Name = '{productName}' AND Product2Id in (SELECT Product__c FROM Product_Junction__c WHERE Product_Line__r.Name = '{productLine}')"
#                 priceBookResult = AccountOpportunityCreation.sf.query_all(
#                     query=priceBookSOQL)
#                 priceBookData = priceBookResult['records']
#                 if priceBookData[0]['IsActive'] != False:
#                     priceBookId = priceBookData[0]['Id']
#                     Messages.write_message(f"Pricebook Entry ID: {priceBookId}")
#                     print("Pricebook Entry ID: ", priceBookId)
#                     data_store.spec["PRICE_BOOK_ID"] = priceBookId

#                     if productLine in ('InStore', 'InStore-Canada'):
#                         chargeDetailsSOQL = f"SELECT Id, Name, ProductLIne__c, Product__C, Charge_Type_2__c, Charge_Type_2__r.Commissionable__c FROM Pricing_Detail__c where ProductLIne__c = 'Instore' and product__c = '{productName}' and CurrencyIsoCode = '{currencyCode}' and isActive__c = true and IsCurrent__c = true"
#                     elif productLine in ('Checkout 51'):
#                         chargeDetailsSOQL = f"SELECT Id, Name, ProductLIne__c, Product__C, Charge_Type_2__c, Charge_Type_2__r.Commissionable__c FROM Pricing_Detail__c where ProductLIne__c = '{productLine}' and FreedomOrNot__c != 'Freedom' and product__c = '{productName}' and CurrencyIsoCode = '{currencyCode}' and isActive__c = true and IsCurrent__c = true"
#                     elif productLine in ('Digital','Digital- Canada'):
#                         chargeDetailsSOQL = f"SELECT Id, Name, ProductLIne__c, Product__C, Charge_Type_2__c, Charge_Type_2__r.Commissionable__c FROM Pricing_Detail__c where ProductLIne__c = '{productLine}' and product__c = '{productName}' and CurrencyIsoCode = '{currencyCode}' and isActive__c = true and FreedomOrNot__c not in ('Freedom') and IsCurrent__c = true"
#                     else:
#                         chargeDetailsSOQL = f"SELECT Id, Name, ProductLIne__c, Product__C, Charge_Type_2__c, Charge_Type_2__r.Commissionable__c FROM Pricing_Detail__c where ProductLIne__c = '{productLine}' and product__c = '{productName}' and CurrencyIsoCode = '{currencyCode}' and isActive__c = true and IsCurrent__c = true"
#                     chargeDetailsResult = AccountOpportunityCreation.sf.query_all(
#                         query=chargeDetailsSOQL)
#                     chargeDetailsData = chargeDetailsResult['records']
#                     print("\n", chargeDetailsSOQL, "\n")
#                     if len(chargeDetailsData) > 2:
#                         chargeTypeMap = {}
#                         chargeTypeMap.clear()
#                         for chargeDetails in chargeDetailsData:
#                             chargeTypeMap[chargeDetails['Name']] = {"Id": chargeDetails['Id'], "ChargeType": chargeDetails[
#                                 'Charge_Type_2__c'], 'Commissionable': chargeDetails['Charge_Type_2__r']['Commissionable__c']}
#                         # print("\nCharge Types\n", chargeTypeMap)
#                         [print(key, value) for key, value in chargeTypeMap.items()]
#                         Messages.write_message(f"Charge Types MAP \n{chargeTypeMap}")
#                         IsActive = True
#             lineItemsMap = {}
#             # pdb.set_trace()
#             for childOpp in childOpps:
#                 totalAmount = 0
#                 for cnt in range(1, 5):
#                     quantity = randint(1, 100)
#                     salesPrice = round(random.uniform(1.1, 9.9), 2)
#                     totalPrice = round((quantity * salesPrice), 2)
#                     totalAmount = totalAmount + totalPrice

#                     chargeType = random.choice(list(chargeTypeMap))
#                     # pdb.set_trace()
#                     lineItemsDetails = AccountOpportunityCreation.sf.OpportunityLineItem.create({'Charge_Type__c': chargeTypeMap[chargeType]["ChargeType"], 'Commissionable__c': chargeTypeMap[chargeType]["Commissionable"], 'pricebookentryid': priceBookId,
#                                                                                                  'Pricing_Detail__c': chargeType, 'PricingDetail__c': chargeTypeMap[chargeType]["Id"], 'opportunityId': childOpp, 'Quantity': quantity, 'Sales_price__c': salesPrice, 'TotalPrice': totalPrice})
                    
#                     data_store.spec[f'{cnt}_CHARGE_TYPE'] = chargeTypeMap[chargeType]["ChargeType"]
#                     data_store.spec[f'{cnt}_COMMISSIONABLE'] = chargeTypeMap[chargeType]["Commissionable"]
#                     data_store.spec[f'{cnt}_PRICING_DETAIL_NAME'] = chargeType
#                     data_store.spec[f'{cnt}_PRICING_DETAIL_ID'] = chargeTypeMap[chargeType]["Id"]
#                     data_store.spec[f'{cnt}_QTY'] = quantity
#                     data_store.spec[f'{cnt}_SALES_PRICE'] = salesPrice
#                     data_store.spec[f'{cnt}_SUB_TOTAL'] = totalPrice
#                     data_store.spec[f'{cnt}_TOTAL_AMOUNT'] = totalAmount
                    
#                     if lineItemsDetails["success"] == True:
#                         print(
#                             f"OPP[{childOpp}] Line Item {cnt}: {lineItemsDetails['id']} created..")
#                         Messages.write_message(f"OPP[{childOpp}] Line Item {cnt}: {lineItemsDetails['id']} created..")
#                         lineItemsMap[childOpp +
#                                      str(cnt)] = {cnt: lineItemsDetails['id']}
#             # print("\nLine Items\n", lineItemsMap)
#             print(json.dumps(lineItemsMap, indent=4))
#             usersVerified = []
#             usersNotVerified = []

#             # Verify Opportunities Team
#             accountTeamMap = {}
#             # '{accID}'"
#             accountMembersSOQL = f"SELECT Account__r.Acc_Owner_Terr_Cat__c,Role_In_Territory__c ,TerritoryId__c,Territory_Category__c,User__c,User__r.Name FROM Account_Team__c where Account__c = '{accountID}'"
#             accountMembersResult = AccountOpportunityCreation.sf.query_all(
#                 query=accountMembersSOQL)
#             accountMembersData = accountMembersResult['records']
#             if len(accountMembersData) > 0:
#                 for accountMembers in accountMembersData:
#                     accountTeamMap[accountMembers["User__c"]] = {"Role_In_Territory__c": accountMembers["Role_In_Territory__c"], "TerritoryId__c": accountMembers["TerritoryId__c"],
#                                                                  "Territory_Category__c": accountMembers["Territory_Category__c"], "Name": accountMembers["User__r"]["Name"], "Acc_Owner_Terr_Cat__c": accountMembers["Account__r"]["Acc_Owner_Terr_Cat__c"]}

#             for childOpp in childOpps:
#                 opportunityMembersSOQL = f"SELECT Opportunity__r.RecordType.Name,Id,TeamMemberRole__c,TerritoryId__c,Territory_Category__c,User__c,User__r.Name FROM Opportunity_Team__c where Opportunity__c = '{childOpp}'"
#                 opportunityMembersResult = AccountOpportunityCreation.sf.query_all(
#                     query=opportunityMembersSOQL)
#                 opportunityMembersData = opportunityMembersResult['records']
#                 if len(opportunityMembersData) > 0:
#                     for opportunityMembers in opportunityMembersData:
#                         isMemberVerified = False
#                         for accountMembers in accountTeamMap:
#                             if not isMemberVerified:
#                                 #                         pdb.set_trace()
#                                 print("\nVerifying Users", opportunityMembers['User__r']['Name'],
#                                       "\t", accountTeamMap[opportunityMembers['User__c']]['Name'])
#                 #                 if (accountTeamMap['Role_In_Territory__c'] == territoryMembers['RoleInTerritory2']) and (accountTeamMap['TerritoryId__c'] == territoryMembers['Territory2Id']) and (accountTeamMap['User__c'] == territoryMembers['UserId']):
#                                 if opportunityMembers["User__r"]["Name"] == accountTeamMap[opportunityMembers["User__c"]]["Name"]:
#                                     print(
#                                         "\nMember Verified: ", accountTeamMap[opportunityMembers['User__c']]['Name'], "\t", opportunityMembers['User__r']['Name'])
#                                     if accountTeamMap[opportunityMembers['User__c']]['Acc_Owner_Terr_Cat__c'] == 'Core':
#                                         if opportunityMembers['Opportunity__r']['RecordType']['Name'] in ('FSI', 'InStore', 'SSMG'):
#                                             if opportunityMembers['Territory_Category__c'] in ('Core') and opportunityMembers['TeamMemberRole__c'] in ('Primary'):
#                                                 print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r'][
#                                                     'RecordType']['Name'], "\t", opportunityMembers['TeamMemberRole__c'])
#                                             elif opportunityMembers['Territory_Category__c'] in ('Checkout_51', 'Digital', 'Merchandising', 'SSD') and opportunityMembers['TeamMemberRole__c'] in ('Specialty', 'Integrated'):
#                                                 print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r'][
#                                                     'RecordType']['Name'], "\t", opportunityMembers['TeamMemberRole__c'])
#                                     elif accountTeamMap[opportunityMembers['User__c']]['Acc_Owner_Terr_Cat__c'] == 'Checkout_51':
#                                         if opportunityMembers['Opportunity__r']['RecordType']['Name'] in ('Checkout 51'):
#                                             if opportunityMembers['Territory_Category__c'] in ('Checkout_51') and opportunityMembers['TeamMemberRole__c'] in ('Primary'):
#                                                 print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r'][
#                                                     'RecordType']['Name'], "\t", opportunityMembers['TeamMemberRole__c'])
#                                     elif accountTeamMap[opportunityMembers['User__c']]['Acc_Owner_Terr_Cat__c'] == 'Digital':
#                                         if opportunityMembers['Opportunity__r']['RecordType']['Name'] in ('Digital', 'Digital- Canada'):
#                                             if opportunityMembers['Territory_Category__c'] in ('Digital') and opportunityMembers['TeamMemberRole__c'] in ('Primary'):
#                                                 print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r'][
#                                                     'RecordType']['Name'], "\t", opportunityMembers['TeamMemberRole__c'])
#                                     elif accountTeamMap[opportunityMembers['User__c']]['Acc_Owner_Terr_Cat__c'] == 'Merchandising':
#                                         if opportunityMembers['Opportunity__r']['RecordType']['Name'] in ('Merchandising'):
#                                             if opportunityMembers['Territory_Category__c'] in ('Merchandising') and opportunityMembers['TeamMemberRole__c'] in ('Primary'):
#                                                 print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r'][
#                                                     'RecordType']['Name'], "\t", opportunityMembers['TeamMemberRole__c'])
#                                     elif accountTeamMap[opportunityMembers['User__c']]['Acc_Owner_Terr_Cat__c'] == 'NTL':
#                                         if opportunityMembers['Territory_Category__c'] in ('NTL') and opportunityMembers['TeamMemberRole__c'] in ('Primary'):
#                                             print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r'][
#                                                 'RecordType']['Name'], "\t", opportunityMembers['TeamMemberRole__c'])
#                                     elif accountTeamMap[opportunityMembers['User__c']]['Acc_Owner_Terr_Cat__c'] == 'SM':
#                                         if opportunityMembers['Territory_Category__c'] in ('SM') and opportunityMembers['TeamMemberRole__c'] in ('Primary'):
#                                             print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r'][
#                                                 'RecordType']['Name'], "\t", opportunityMembers['TeamMemberRole__c'])
#                                     elif accountTeamMap[opportunityMembers['User__c']]['Acc_Owner_Terr_Cat__c'] == 'SSD':
#                                         if opportunityMembers['Opportunity__r']['RecordType']['Name'] in ('SmartSource Direct', 'SmartSource Direct- Canada'):
#                                             if opportunityMembers['Territory_Category__c'] in ('SSD') and opportunityMembers['TeamMemberRole__c'] in ('Primary'):
#                                                 print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r'][
#                                                     'RecordType']['Name'], "\t", opportunityMembers['TeamMemberRole__c'])
#                                     elif accountTeamMap[opportunityMembers['User__c']]['Acc_Owner_Terr_Cat__c'] == 'SSMG':
#                                         if opportunityMembers['Opportunity__r']['RecordType']['Name'] in ('SSMG'):
#                                             if opportunityMembers['Territory_Category__c'] in ('SSMG') and opportunityMembers['TeamMemberRole__c'] in ('Primary'):
#                                                 print("\n Territory Category Verified: ", opportunityMembers['Opportunity__r']['RecordType'][
#                                                     'Name'], "\t", accountTeamMap[opportunityMembers['User__c']]['Territory_Category__c'])

#                                     if opportunityMembers['User__r']['Name'] not in usersVerified:
#                                         usersVerified.append(
#                                             opportunityMembers['User__r']['Name'])

#                                     if opportunityMembers['User__r']['Name'] in usersNotVerified:
#                                         usersNotVerified.remove(
#                                             opportunityMembers['User__r']['Name'])
#                                     isMemberVerified = True

#                                 else:
#                                     #                 print("Member Not Verified: ", accountTeamMap['User__r']['Name'], "\t", accountTeamMap['Role_In_Territory__c'])
#                                     if opportunityMembers['User__r']['Name'] not in usersNotVerified:
#                                         usersNotVerified.append(
#                                             opportunityMembers['User__r']['Name'])

#                                     if opportunityMembers['User__r']['Name'] in usersVerified:
#                                         usersVerified.remove(
#                                             opportunityMembers['User__r']['Name'])
#                                     isMemberVerified = False

#                             print("\nMembers Verified: ", usersVerified)
#                             print("\nMembers not verified: ", usersNotVerified)

#             print("\nUS Product Lines\n", usProductLineMap)
#             jsonData = json.dumps(usProductLineMap)
#             with open("OpportunitiesDetails.json", 'w') as f:
#                 f.write(jsonData)
#             print(
#                 f"{productLine} completed---------------------------------------------------------------------------\n")

