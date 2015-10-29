/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import "SnippetsManager.h"
#import "NetworkManager.h"
#import "Operation.h"
#import "AuthenticationManager.h"
#import "MultiformObject.h"
#import "SnippetsManagerConf.h"


@interface SnippetsManager()

@end

@implementation SnippetsManager

- (instancetype)init{
    self = [super init];
    if(self){
        // Initialize sections
        _sections = @[@"Users",
                      @"Groups",
                      @"Contacts"];
        
        // Initialize operations
        _operationsArray = [NSMutableArray new];
        
        
        // Section - Users
        NSMutableArray *usersArray = [NSMutableArray new];
        
        [usersArray addObject:[self getUsersInTenant]];
        [usersArray addObject:[self getSelectUsersInTenant]];
        [usersArray addObject:[self createNewUser]];
        
        [usersArray addObject:[self getUserProfile]];
        [usersArray addObject:[self getPropertiesUserProfile]];
        
        [usersArray addObject:[self getUserDrive]];
        [usersArray addObject:[self getUserEvents]];
        [usersArray addObject:[self addNewCalendarEvent]];  // create new event
        
        [usersArray addObject:[self updateCalendarEvent]];
        [usersArray addObject:[self deleteCalendarEvent]];
        
        [usersArray addObject:[self getUserMessages]];
        [usersArray addObject:[self createAndSendMessage]];
        
        [usersArray addObject:[self getUserManager]];
        [usersArray addObject:[self getUserReports]];
        [usersArray addObject:[self getUserPhoto]];
        [usersArray addObject:[self getUserGroupMembership]];
        [usersArray addObject:[self getUserFiles]];
        [usersArray addObject:[self createNewFile]];
        [usersArray addObject:[self createNewFolder]];
        [usersArray addObject:[self downloadFile]];
        [usersArray addObject:[self updateFileContents]];
        [usersArray addObject:[self deleteFile]];
        [usersArray addObject:[self copyFile]];
        [usersArray addObject:[self updateFileMetadata]];
        
        
        // Section 2 - Groups
        NSMutableArray *groupsArray = [NSMutableArray new];
        [groupsArray addObject:[self getGroupsInTenant]];
        [groupsArray addObject:[self addNewSecurityGroup]];
        [groupsArray addObject:[self getSpecificGroup]];
        [groupsArray addObject:[self updateGroup]];
        [groupsArray addObject:[self deleteGroup]];
        [groupsArray addObject:[self getGroupMembers]];
        [groupsArray addObject:[self getGroupOwners]];
        
        // Section 3 - Contacts
        NSMutableArray *contactsArray = [NSMutableArray new];
        [contactsArray addObject:[self getContactsInTenant]];
        
        // Add sections to the array
        [_operationsArray addObject:usersArray];
        [_operationsArray addObject:groupsArray];
        [_operationsArray addObject:contactsArray];
        
        
    }
    return self;
}


#pragma mark - Users

//Gets all user files.
- (Operation *) getUserFiles {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user files"
                                                          urlString:[self createURLString:@"/me/drive/root/children"]
                                                      operationType:OperationGet
                                                        description:@"Returns all of the user's files"
                                                  documentationLink:@"https://dev.onedrive.com/drives/get.htm"
                                                             params:nil
                                                       paramsSource:nil];

    return operation;
    
}

// Creates a text file in the user's root directory.
- (Operation *) createNewFile {
    NSString *fileName = @"NewFile.txt";
    Operation *operation = [[Operation alloc] initWithOperationName:@"PUT: Create a text file"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/me/drive/root/children/%@/content", fileName]]
                                                      operationType:OperationPut
                                                       customHeader:@{@"content-type":@"text/plain"}
                                                         customBody:@"Test file"
                                                        description:@"Creates a text file in the user's root directory. The file name NewFile.txt is editable in API URL. "
                                                  documentationLink:@"https://dev.onedrive.com/items/create.htm"
                                                             params:nil
                                                       paramsSource:nil];
    
    return operation;
}

// Creates a folder in the user's root directory - newFolderData.json
- (Operation *) createNewFolder {
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"newFolderData" ofType:@"json"];
    NSMutableString *payload = [NSMutableString stringWithString:[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil]];
    
    Operation *operation = [[Operation alloc] initWithOperationName:@"POST: Create a new folder"
                                                          urlString:[self createURLString:@"/me/drive/root/children"]
                                                      operationType:OperationPostCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Creates a folder in the user's root directory."
                                                  documentationLink:@"https://dev.onedrive.com/items/create.htm"
                                                             params:@{ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsPostDataKey:@(ParamsSourcePostData)}];
    
    return operation;
}


//Downloads the content of an existing file.
- (Operation *) downloadFile {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Download file content"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/me/drive/items/{%@}", ParamsFileIDKey]]
                                                      operationType:OperationGet
                                                        description:@"Downloads the content of an existing file"
                                                  documentationLink:@"https://dev.onedrive.com/items/download.htm"
                                                             params:@{ParamsFileIDKey:@""}
                                                       paramsSource:@{ParamsFileIDKey:@(ParamsSourceGetFiles)}];
    
    return operation;
    
}


// Update file contents, name of file -patchFileData.json
- (Operation *) updateFileContents {
    Operation *operation = [[Operation alloc] initWithOperationName:@"PATCH: Updates file metadata"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/me/drive/items/{%@}/content", ParamsFileIDKey]]
                                                      operationType:OperationPut
                                                       customHeader:@{@"content-type":@"text/plain"}
                                                         customBody:@"Updated text"
                                                        description:@"Updates the content of the selected file."
                                                  documentationLink:@"https://dev.onedrive.com/items/update.htm"
                                                             params:@{ParamsFileIDKey:@""}
                                                       paramsSource:@{ParamsFileIDKey:@(ParamsSourceGetFiles)}];
    
    return operation;
}

// Deletes a file in the user's root directory.
-(Operation *) deleteFile {
    Operation *operation = [[Operation alloc] initWithOperationName:@"DELETE: Delete a file"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/me/drive/items/{%@}", ParamsFileIDKey]]
                                                      operationType:OperationDelete
                                                        description:@"Deletes a file in the user's root directory."
                                                  documentationLink:@"https://dev.onedrive.com/items/delete.htm"
                                                             params:@{ParamsFileIDKey: @""}
                                                       paramsSource:@{ParamsFileIDKey: @(ParamsSourceGetFiles)}];
    return operation;
}


// Copies a file in the user's root directory
- (Operation *) copyFile {
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"copyFileData" ofType:@"json"];
    NSString *payload = [NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil];
    
    
    Operation *operation = [[Operation alloc] initWithOperationName:@"POST: Copy a text file"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/me/drive/items/{%@}/microsoft.graph.copy", ParamsFileIDKey]]
                                                      operationType:OperationPostCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Copies a file in the user's root directory."
                                                  documentationLink:@"https://dev.onedrive.com/items/copy.htm"
                                                             params:@{ParamsFileIDKey:@"",
                                                                      ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsFileIDKey:@(ParamsSourceGetFiles),
                                                                      ParamsPostDataKey:@(ParamsSourcePostData)}];
    
    return operation;
}


//Updates file metadata in the user's root directory - patchMetadataFile.json
- (Operation *) updateFileMetadata {
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"patchMetadataFile" ofType:@"json"];
    NSMutableString *payload = [NSMutableString stringWithString:[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil]];
    
    Operation *operation = [[Operation alloc] initWithOperationName:@"PATCH: Updates file metadata"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/me/drive/items/{%@}", ParamsFileIDKey]]
                                                      operationType:OperationPatchCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Renames a file's metadata in the user's root directory."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Event"
                                                             params:@{ParamsFileIDKey:@"",
                                                                      ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsFileIDKey:@(ParamsSourceGetFiles),
                                                                      ParamsPostDataKey:@(ParamsSourcePostData)}];
    
    return operation;
}



//Returns all of the users in your tenant's directory.
- (Operation *) getUsersInTenant {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get users in tenant"
                                                          urlString:[self createURLString:@"/myOrganization/users"]
                                                      operationType:OperationGet
                                                        description:@"Returns all of the users in your tenant's directory"
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_User"
                                                             params:nil
                                                       paramsSource:nil];
    
    return operation;
}

//Returns select users in your tenant's directory.
- (Operation *) getSelectUsersInTenant {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get select users in a tenant"
                                                          urlString:[self createURLString:@"/myOrganization/users"]
                                                      operationType:OperationGet
                                                        description:@"Returns all of the users in your tenant's directory who are from the United States, using $filter."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_User"
                                                             params:@{@"$filter":@"country eq 'United States'"}
                                                       paramsSource:@{@"$filter":@(ParamsSourceTextEdit)}];
    return operation;
}

// Create new user - newUserData.json
- (Operation *) createNewUser {
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"newUserData" ofType:@"json"];
    
    NSString *userId = [[NSProcessInfo processInfo] globallyUniqueString];
    NSString *domainName = [[[[AuthenticationManager sharedInstance] userID] componentsSeparatedByString:@"@"] lastObject];
    NSString *payload = [[[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil]
                          stringByReplacingOccurrencesOfString:@"<USER>" withString:userId]
                         stringByReplacingOccurrencesOfString:@"<DOMAIN>" withString:domainName];
    
    Operation *operation = [[Operation alloc] initWithOperationName:@"POST: Create a new user"
                                                          urlString:[self createURLString:@"/myOrganization/users"]
                                                      operationType:OperationPostCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Create a new user and adds them to the user collection."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_User"
                                                             params:@{ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsPostDataKey:@(ParamsSourcePostData)}];
    
    operation.isAdminRequired = YES;
    return operation;
}


//Returns the user's profile.
- (Operation *) getUserProfile {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user's profile"
                                                          urlString:[self createURLString:@"/me"]
                                                      operationType:OperationGet
                                                        description:@"Returns all of the users in your tenant's directory who are from the United States, using $filter."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_User"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}




//Returns select information about the signed-in user from Azure Active Directory.
-(Operation *) getPropertiesUserProfile {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get select properties of user's profile"
                                                          urlString:[self createURLString:@"/me"]
                                                      operationType:OperationGet
                                                        description:@"Returns select information about the signed-in user from Azure Active Directory."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_User"
                                                             params:@{@"$select":@"AboutMe,Responsibilities,Tags"}
                                                       paramsSource:@{@"$filter":@(ParamsSourceTextEdit)}];
    return operation;
}


//Gets the signed-in user's drive from OneDrive for Business.
- (Operation *) getUserDrive {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user's drive"
                                                          urlString:[self createURLString:@"/me/drive"]
                                                      operationType:OperationGet
                                                        description:@"Gets the signed-in user's drive from OneDrive for Business."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Drive"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}


//Gets the signed-in user's events from Office 365.
- (Operation *) getUserEvents {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user's events"
                                                          urlString:[self createURLString:@"/me/events"]
                                                      operationType:OperationGet
                                                        description:@"Gets the signed-in user's events from Office 365."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Event"
                                                             params:@{@"$select":@"id, Subject"}
                                                       paramsSource:@{@"$select":@(ParamsSourceTextEdit)}];
    return operation;
}


// Creates and adds an event to the signed-in user's calendar - eventData.json
- (Operation *) addNewCalendarEvent {
    NSDateFormatter *dateFormatter = [[NSDateFormatter alloc] init];
    [dateFormatter setDateFormat:@"yyyy-MM-dd'T'HH:mm:ss.SSS'Z'"];
    
    NSDate *startDate = [[NSDate date] dateByAddingTimeInterval:60*60*24];
    NSDate *endDate = [[NSDate date] dateByAddingTimeInterval:60*60*24+100];
    
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"eventData" ofType:@"json"];
    NSString *payload = [[[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil] stringByReplacingOccurrencesOfString:@"<STARTDATETIME>" withString:[dateFormatter stringFromDate:startDate]]
                         stringByReplacingOccurrencesOfString:@"<ENDDATETIME>" withString:[dateFormatter stringFromDate:endDate]];
    
    Operation *operation = [[Operation alloc] initWithOperationName:@"POST: Add a new event"
                                                          urlString:[self createURLString:@"/me/events"]
                                                      operationType:OperationPostCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Creates and adds an event to the signed-in user's calendar."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Event"
                                                             params:@{ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsPostDataKey:@(ParamsSourcePostData)}];
    return operation;
}

// Updates the event from user's calendar - eventDataPatch.json
- (Operation *) updateCalendarEvent {
    
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"eventDataPatch" ofType:@"json"];
    NSMutableString *payload = [NSMutableString stringWithString:[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil]];
    
    //NSString *payload = @"{Subject: 'Weekly Sync', Location: { DisplayName: 'Water cooler'";
    Operation *operation = [[Operation alloc] initWithOperationName:@"PATCH: Update an event"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/me/events/{%@}", ParamsEventIDKey]]
                                                      operationType:OperationPatchCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Updates a signed-in user's calendar."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Event"
                                                             params:@{ParamsEventIDKey:@"",
                                                                      ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsEventIDKey:@(ParamsSourceGetEvents),
                                                                      ParamsPostDataKey:@(ParamsSourcePostData)}];
    
    return operation;
}


// Deletes an event from user's calendar
-(Operation *) deleteCalendarEvent {
    Operation *operation = [[Operation alloc] initWithOperationName:@"DELETE: Delete an event"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/me/events/{%@}", ParamsEventIDKey]]
                                                      operationType:OperationDelete
                                                        description:@"Creates and adds an event to the signed-in user's calendar, then deletes the event."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Event"
                                                             params:@{ParamsEventIDKey: @""}
                                                       paramsSource:@{ParamsEventIDKey: @(ParamsSourceGetEvents)}];
    return operation;
}


//Gets the signed-in user's messages from Office 365.
- (Operation *) getUserMessages {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user's messages"
                                                          urlString:[self createURLString:@"/me/messages"]
                                                      operationType:OperationGet
                                                        description:@"Gets the signed-in user's messages from Office 365."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_relationship_Messages"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}

// Create and send a message as the signed-in user - emailData.json
- (Operation *) createAndSendMessage {
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"emailData" ofType:@"json"];
    
    // replace email address to self
    NSString *payload = [[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil]
                         stringByReplacingOccurrencesOfString:@"<EMAIL>" withString:[[AuthenticationManager sharedInstance] userID]];
    
    
    Operation *operation = [[Operation alloc] initWithOperationName:@"POST: Create and send message to user"
                                                          urlString:[self createURLString:@"/me/sendMail"]
                                                      operationType:OperationPostCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Create and send a message as the signed-in user."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_action_user_sendMail"
                                                             params:@{ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsPostDataKey:@(ParamsSourcePostData)}];
    
    return operation;
}

//GET: Get user's manager
- (Operation *) getUserManager {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user's manager"
                                                          urlString:[self createURLString:@"/me/manager"]
                                                      operationType:OperationGet
                                                        description:@"GET: Get user's manager"
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_relationship_manager"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}

//Gets the signed-in user's direct reports.
- (Operation *) getUserReports {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user's direct reports"
                                                          urlString:[self createURLString:@"/me/directReports"]
                                                      operationType:OperationGet
                                                        description:@"Gets the signed-in user's direct reports."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_relationship_directReports"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}

//Gets the signed-in user's photo.
- (Operation *) getUserPhoto{
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user's photo"
                                                          urlString:[self createURLString:@"/me/Photo"]
                                                      operationType:OperationGet
                                                        description:@"Gets the signed-in user's photo."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_relationship_UserPhoto"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}

//Gets a collection of groups that the signed-in user is a member of.
- (Operation *) getUserGroupMembership {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get user group membership"
                                                          urlString:[self createURLString:@"/me/memberOf"]
                                                      operationType:OperationGet
                                                        description:@"Gets a collection of groups that the signed-in user is a member of."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_relationship_memberOf"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}

#pragma mark - Groups Snippets

//Returns all of the groups in your tenant's directory.
- (Operation *) getGroupsInTenant {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get groups in tenant"
                                                          urlString:[self createURLString:@"/myOrganization/groups"]
                                                      operationType:OperationGet
                                                        description:@"Returns all of the groups in your tenant's directory"
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entitySet_groups"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}

// Add new security group
- (Operation *) addNewSecurityGroup {
    
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"newGroupData" ofType:@"json"];
    NSString *payload = [[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil] stringByReplacingOccurrencesOfString:@"<GROUP>" withString:[[NSProcessInfo processInfo] globallyUniqueString]];
    
    Operation *operation = [[Operation alloc] initWithOperationName:@"POST: Add a new security group"
                                                          urlString:[self createURLString:@"/myOrganization/groups"]
                                                      operationType:OperationPostCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Adds a new security group to the tenant."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entitySet_groups"
                                                             params:@{ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsPostDataKey:@(ParamsSourcePostData)}];
    return operation;
}

// Gets information about a specific group in the tenant by ID.
- (Operation *) getSpecificGroup {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get specific group by ID"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/myOrganization/groups/{%@}", ParamsGroupIDKey]]
                                                      operationType:OperationGet
                                                        description:@"Gets information about a specific group in the tenant by ID."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Group"
                                                             params:@{ParamsGroupIDKey:@""}
                                                       paramsSource:@{ParamsGroupIDKey:@(ParamsSourceGetGroups)}];
    return operation;
}

// Updates the description of the group - patchGroupData.json
- (Operation *) updateGroup {
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"patchGroupData" ofType:@"json"];
    NSMutableString *payload = [NSMutableString stringWithString:[NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil]];
    
    Operation *operation = [[Operation alloc] initWithOperationName:@"PATCH: Updates specific group"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/myOrganization/groups/{%@}", ParamsGroupIDKey]]
                                                      operationType:OperationPatchCustom
                                                       customHeader:@{@"content-type":@"application/json"}
                                                         customBody:payload
                                                        description:@"Updates the description of the group."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Event"
                                                             params:@{ParamsGroupIDKey:@"",
                                                                      ParamsPostDataKey:payload}
                                                       paramsSource:@{ParamsGroupIDKey:@(ParamsSourceGetGroups),
                                                                      ParamsPostDataKey:@(ParamsSourcePostData)}];
    
    return operation;
}

// Deletes a group
- (Operation *) deleteGroup {
    Operation *operation = [[Operation alloc] initWithOperationName:@"DELETE: Delete a group"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/myOrganization/groups/{%@}", ParamsGroupIDKey]]
                                                      operationType:OperationDelete
                                                        description:@"Deletes a group."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entityType_Group"
                                                             params:@{ParamsGroupIDKey:@""}
                                                       paramsSource:@{ParamsGroupIDKey:@(ParamsSourceGetGroups)}];
    return operation;
}

// Gets a specific group's members
- (Operation *) getGroupMembers {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get specific group members"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/myOrganization/groups/{%@}/members", ParamsGroupIDKey]]
                                                      operationType:OperationGet
                                                        description:@"Gets a specific group's members."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_relationship_members"
                                                             params:@{ParamsGroupIDKey:@""}
                                                       paramsSource:@{ParamsGroupIDKey:@(ParamsSourceGetGroups)}];
    return operation;
}

// Gets a specific group's owners
- (Operation *) getGroupOwners {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get specific group owners"
                                                          urlString:[self createURLString:[NSString stringWithFormat:@"/myOrganization/groups/{%@}/owners", ParamsGroupIDKey]]
                                                      operationType:OperationGet
                                                        description:@"Gets a specific group's owners."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_relationship_owners"
                                                             params:@{ParamsGroupIDKey:@""}
                                                       paramsSource:@{ParamsGroupIDKey:@(ParamsSourceGetGroups)}];
    return operation;
}

#pragma mark - Contacts Snippets

//Returns all of the contacts in your tenant's directory.
- (Operation *) getContactsInTenant {
    Operation *operation = [[Operation alloc] initWithOperationName:@"GET: Get contacts in tenant"
                                                          urlString:[self createURLString:@"/myOrganization/contacts"]
                                                      operationType:OperationGet
                                                        description:@"Returns all of the contacts in your tenant's directory."
                                                  documentationLink:@"https://msdn.microsoft.com/office/office365/HowTo/office-365-unified-api-reference#msg_ref_entitySet_contacts"
                                                             params:nil
                                                       paramsSource:nil];
    return operation;
}

#pragma mark - helper
- (NSString *) createURLString:(NSString *)path {
    NSMutableString *urlString = [NSMutableString new];
    [urlString appendString:[SnippetsManagerConf protocol]];
    [urlString appendString:@"://"];
    [urlString appendString:[SnippetsManagerConf hostName_beta]];
    [urlString appendString:path];
    return urlString;
}

@end


// *********************************************************
//
// O365-iOS-Unified-API-Snippets, https://github.com/OfficeDev/O365-iOS-Unified-API-Snippets
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
