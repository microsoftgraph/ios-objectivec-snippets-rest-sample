/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import <Foundation/Foundation.h>

typedef enum{
    OperationPost,
    OperationPostCustom,
    OperationGet,
    OperationDelete,
    OperationPostMultiPart,
    OperationPatch,
    OperationPatchCustom
} OperationType;

typedef enum{
    ParamsSourceTextEdit = 1,
    ParamsSourceGetEvents = 2,
    ParamsSourceGetGroups = 3,
    ParamsSourcePostData = 4
} ParamsSourceType;

extern NSString* const ParamsEventIDKey;
extern NSString* const ParamsGroupIDKey;

extern NSString* const ParamsPostDataKey;

@interface Operation : NSObject

@property (nonatomic, strong) NSString *operationName;
@property (nonatomic, strong) NSString *operationURLString;

@property (nonatomic, strong) NSString *descriptionString;
@property (nonatomic, strong) NSString *documentationLinkString;

@property (nonatomic, assign) OperationType operationType;

@property (nonatomic, strong) NSDictionary *customHeader;
@property (nonatomic, strong) NSString *customBody;

@property (nonatomic, strong) NSDictionary *params;
@property (nonatomic, strong) NSDictionary *paramsSource;

@property (nonatomic, strong) NSString *customResponseType;

@property (nonatomic, strong) NSArray *multiPartObjectArray;

@property (nonatomic, assign) BOOL isAdminRequired;

- (instancetype) initWithOperationName:(NSString *)operationName
                             urlString:(NSString *)urlString
                         operationType:(OperationType)operationType
                           description:(NSString *)description
                     documentationLink:(NSString *)documentationLink
                                params:(NSDictionary *)params
                          paramsSource:(NSDictionary *)paramsSource;

- (instancetype) initWithOperationName:(NSString *)operationName
                             urlString:(NSString *)urlString
                         operationType:(OperationType)operationType
                          customHeader:(NSDictionary *)customHeader
                            customBody:(NSString *)customBody
                           description:(NSString *)description
                     documentationLink:(NSString *)documentationLink
                                params:(NSDictionary *)params
                          paramsSource:(NSDictionary *)paramsSource;

- (instancetype) initWithOperationName:(NSString *)operationName
                             urlString:(NSString *)urlString
                         operationType:(OperationType)operationType
                           description:(NSString *)description
                     documentationLink:(NSString *)documentationLink
                      multiPartObjects:(NSArray *)multiPartObjects;

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
