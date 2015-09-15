/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import <Foundation/Foundation.h>
#import "MultiformObject.h"

@interface NetworkManager : NSObject

// GET
+ (void)get:(NSString*)path
queryParams:(NSDictionary*)queryParams
    success:(void (^)(id responseHeader, id responseObject))success
    failure:(void (^)(id responseObject))failure;

// GET with custom response type
// For example, acceptable response type @"text/html"
+ (void)get:(NSString*)path
    queryParams:(NSDictionary*)queryParams
customResponseType:(NSString*)responseType
    success:(void (^)(id responseHeader, id responseObject))success
    failure:(void (^)(id responseObject))failure;

// POST using query params in a dictionary
// This will create json payload in the request
+ (void)post:(NSString*)path
 queryParams:(NSDictionary*)queryParams
     success:(void (^)(id responseHeader, id responseObject))success
     failure:(void (^)(id responseObject))failure;

// POST using custom header and body
+ (void)post:(NSString*)path
customHeader:(NSDictionary*)customHeader
  customBody:(NSString*)bodyString
     success:(void (^)(id responseHeader, id responseObject))success
     failure:(void (^)(id responseObject))failure;

// DELETE
+ (void)deleteOperation:(NSString*)path queryParams:(NSDictionary*)queryParams
     success:(void (^)(id responseHeader, id responseObject))success
     failure:(void (^)(id responseObject))failure;


// PATCH using query params in a dictionary
// This will create json payload in the request
+ (void)patchOperation:(NSString*)path
           queryParams:(NSDictionary*)queryParams
               success:(void (^)(id responseHeader, id responseObject))success
               failure:(void (^)(id responseObject))failure;


// PATCH using custom header and body
+ (void)patchOperation:(NSString*)path
          customHeader:(NSDictionary*)customHeader
            customBody:(NSString*)bodyString
               success:(void (^)(id responseHeader, id responseObject))success
               failure:(void (^)(id responseObject))failure;

// POST multipart form
+ (void)postWithMultipartForm:(NSString*)path
                  queryParams:(NSDictionary*)queryParams
             multiformObjects:(NSArray*)multiformObjects
                      success:(void (^)(id responseHeader, id responseObject))success
                      failure:(void (^)(id responseObject))failure;
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
