/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import "Operation.h"

@implementation Operation

NSString* const ParamsEventIDKey = @"eventID";
NSString* const ParamsPostDataKey = @"Post data";
NSString* const ParamsGroupIDKey = @"groupID";

- (instancetype) initWithOperationName:(NSString *)operationName
                             urlString:(NSString *)urlString
                         operationType:(OperationType)operationType
                           description:(NSString *)description
                     documentationLink:(NSString *)documentationLink
                                params:(NSDictionary *)params
                          paramsSource:(NSDictionary *)paramsSource {
    self = [super init];
    if(self){
        _operationName = operationName;
        _operationURLString = urlString;
        _operationType = operationType;
        _documentationLinkString = documentationLink;
        _descriptionString = description;
        _params = params;
        _paramsSource = paramsSource;
        _customResponseType = nil;
    }
    
    return self;
}

- (instancetype) initWithOperationName:(NSString *)operationName
                             urlString:(NSString *)urlString
                         operationType:(OperationType)operationType
                          customHeader:(NSDictionary *)customHeader
                            customBody:(NSString *)customBody
                           description:(NSString *)description
                     documentationLink:(NSString *)documentationLink
                                params:(NSDictionary *)params
                          paramsSource:(NSDictionary *)paramsSource {
    self = [super init];
    if(self){
        _operationName = operationName;
        _operationURLString = urlString;
        _operationType = operationType;
        _customHeader = customHeader;
        _customBody = customBody;
        _documentationLinkString = documentationLink;
        _descriptionString = description;
        _params = params;
        _paramsSource = paramsSource;
        _customResponseType = nil;
    }
    
    return self;
}

- (instancetype) initWithOperationName:(NSString *)operationName
                             urlString:(NSString *)urlString
                         operationType:(OperationType)operationType
                           description:(NSString *)description
                     documentationLink:(NSString *)documentationLink
                      multiPartObjects:(NSArray *)multiPartObjects {
    self = [super init];
    if(self){
        _operationName = operationName;
        _operationURLString = urlString;
        _operationType = operationType;
        _documentationLinkString = documentationLink;
        _descriptionString = description;
        _multiPartObjectArray = multiPartObjects;
        _customResponseType = nil;
    }
    return self;
    
    
}
- (id)copy {
    Operation *obj = [[Operation alloc] initWithOperationName:self.operationName
                                                    urlString:self.operationURLString
                                                operationType:self.operationType
                                                 customHeader:self.customHeader
                                                   customBody:self.customBody
                                                  description:self.descriptionString
                                            documentationLink:self.documentationLinkString
                                                       params:self.params
                                                 paramsSource:self.paramsSource];
    obj.customResponseType = self.customResponseType;
    obj.multiPartObjectArray = self.multiPartObjectArray;
    obj.isAdminRequired = self.isAdminRequired;
    
    return obj;
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
