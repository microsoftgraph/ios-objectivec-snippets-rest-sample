/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import "NetworkManager.h"
#import <AFNetworking.h>
#import "AuthenticationManager.h"
#import "Operation.h"
#import "AFHTTPResponseSerializerHTML.h"

@interface NetworkManager ()

@property (nonatomic, strong) NSString *protocol;
@property (nonatomic, strong) NSString *host;
@property (nonatomic, strong) NSArray *keysToRemove;
@end

@implementation NetworkManager

// These keys will be removed from the request parameters
+ (NSArray*) keysToRemove{
    static NSArray *_keysToRemove;
    if(_keysToRemove == nil)
        _keysToRemove = @[ParamsEventIDKey, ParamsGroupIDKey];
    return _keysToRemove;
}

// Helper to remove certains keys from the NSDictionary
+ (NSDictionary*) paramsRemove:(NSArray*)keys from:(NSDictionary*)params{
    NSMutableDictionary *newParams = [NSMutableDictionary dictionaryWithDictionary:params];
    [newParams removeObjectsForKeys:keys];
    return newParams;
}

+ (void)get:(NSString*)path
queryParams:(NSDictionary*)queryParams
    success:(void (^)(id responseHeader, id responseObject))success
    failure:(void (^)(id responseObject))failure{
    
    [self get:path queryParams:queryParams
customResponseType:nil
      success:^(id responseHeader, id responseObject) {
          success(responseHeader, responseObject);
      } failure:^(id responseObject) {
          failure(responseObject);
      }];
}

+ (void)get:(NSString*)path
queryParams:(NSDictionary*)queryParams
customResponseType:(NSString*)responseType
    success:(void (^)(id responseHeader, id responseObject))success
    failure:(void (^)(id responseObject))failure{

    [[AuthenticationManager sharedInstance] checkAndRefreshToken:^(ADAuthenticationError *error) {
        if(error){
            failure(error);
            return;
        }

        AFHTTPRequestOperationManager *manager = [AFHTTPRequestOperationManager manager];
        
        [manager.requestSerializer setValue:[NSString stringWithFormat:@"Bearer %@", [[AuthenticationManager sharedInstance] accessToken]]
                         forHTTPHeaderField:@"Authorization"];
        
        if(responseType){
            manager.responseSerializer = [[AFHTTPResponseSerializerHTML alloc] init];
            manager.responseSerializer.acceptableContentTypes = [NSSet setWithObject:responseType];
        }
        
        [manager GET:path
          parameters:[self paramsRemove:self.keysToRemove from:queryParams]
             success:^(AFHTTPRequestOperation *operation, id responseObject) {
                 success(operation.response.allHeaderFields, responseObject);
             } failure:^(AFHTTPRequestOperation *operation, NSError *error) {
                 NSMutableDictionary *userInfo = [NSMutableDictionary dictionaryWithDictionary:error.userInfo];
                 
                 if(![operation.responseString isEqual:[NSNull null]])
                     [userInfo setObject:operation.responseString forKey:@"responseString"];
                 
                 NSError *newError = [NSError errorWithDomain:error.domain
                                                         code:error.code
                                                     userInfo:userInfo];
                 
                 
                 failure(newError);
             }];
    }];
}


+ (void)post:(NSString*)path
 queryParams:(NSDictionary*)queryParams
     success:(void (^)(id responseHeader, id responseObject))success
     failure:(void (^)(id responseObject))failure{
    [[AuthenticationManager sharedInstance] checkAndRefreshToken:^(ADAuthenticationError *error) {
        if(error){
            failure(error);
            return;
        }
        AFHTTPRequestOperationManager *manager = [AFHTTPRequestOperationManager manager];
        
        AFJSONRequestSerializer *serializer = [AFJSONRequestSerializer serializer];
        [serializer setValue:[NSString stringWithFormat:@"Bearer %@", [[AuthenticationManager sharedInstance] accessToken]]
          forHTTPHeaderField:@"Authorization"];
        manager.requestSerializer = serializer;
        
        [manager POST:path
           parameters:[self paramsRemove:self.keysToRemove from:queryParams]
              success:^(AFHTTPRequestOperation *operation, id responseObject) {
                  success(operation.response.allHeaderFields, responseObject);
              } failure:^(AFHTTPRequestOperation *operation, NSError *error) {
                  NSMutableDictionary *userInfo = [NSMutableDictionary dictionaryWithDictionary:error.userInfo];
                  
                  if(![operation.responseString isEqual:[NSNull null]])
                      [userInfo setObject:operation.responseString forKey:@"responseString"];
                  
                  NSError *newError = [NSError errorWithDomain:error.domain
                                                          code:error.code
                                                      userInfo:userInfo];
                  
                  
                  failure(newError);
              }];
    }];
}

+ (void)deleteOperation:(NSString*)path queryParams:(NSDictionary*)queryParams
                success:(void (^)(id responseHeader, id responseObject))success
                failure:(void (^)(id responseObject))failure{
    [[AuthenticationManager sharedInstance] checkAndRefreshToken:^(ADAuthenticationError *error) {
        if(error){
            failure(error);
            return;
        }
        AFHTTPRequestOperationManager *manager = [AFHTTPRequestOperationManager manager];
        
        [manager.requestSerializer setValue:[NSString stringWithFormat:@"Bearer %@", [[AuthenticationManager sharedInstance] accessToken]]
                         forHTTPHeaderField:@"Authorization"];
        
        [manager DELETE:path
             parameters:[self paramsRemove:self.keysToRemove from:queryParams]
                success:^(AFHTTPRequestOperation *operation, id responseObject) {
                    success(operation.response.allHeaderFields, responseObject);
                } failure:^(AFHTTPRequestOperation *operation, NSError *error) {
                    NSMutableDictionary *userInfo = [NSMutableDictionary dictionaryWithDictionary:error.userInfo];
                    
                    if(![operation.responseString isEqual:[NSNull null]])
                        [userInfo setObject:operation.responseString forKey:@"responseString"];
                    
                    NSError *newError = [NSError errorWithDomain:error.domain
                                                            code:error.code
                                                        userInfo:userInfo];
                    
                    
                    failure(newError);
                }];
    }];
}

+ (void)patchOperation:(NSString*)path queryParams:(NSDictionary*)queryParams
               success:(void (^)(id responseHeader, id responseObject))success
               failure:(void (^)(id responseObject))failure{
    [[AuthenticationManager sharedInstance] checkAndRefreshToken:^(ADAuthenticationError *error) {
        if(error){
            failure(error);
            return;
        }
        AFHTTPRequestOperationManager *manager = [AFHTTPRequestOperationManager manager];
        
        AFJSONRequestSerializer *serializer = [AFJSONRequestSerializer serializer];
        [serializer setValue:[NSString stringWithFormat:@"Bearer %@", [[AuthenticationManager sharedInstance] accessToken]]
          forHTTPHeaderField:@"Authorization"];
        [serializer setValue:@"" forHTTPHeaderField:@"Accept-Encoding"];
        
        manager.requestSerializer = serializer;
        
        [manager PATCH:path
            parameters:[NSArray arrayWithObject:[self paramsRemove:self.keysToRemove from:queryParams]]
               success:^(AFHTTPRequestOperation *operation, id responseObject) {
                   success(operation.response.allHeaderFields, responseObject);
               } failure:^(AFHTTPRequestOperation *operation, NSError *error) {
                   NSMutableDictionary *userInfo = [NSMutableDictionary dictionaryWithDictionary:error.userInfo];
                   
                   if(![operation.responseString isEqual:[NSNull null]])
                       [userInfo setObject:operation.responseString forKey:@"responseString"];
                   
                   NSError *newError = [NSError errorWithDomain:error.domain
                                                           code:error.code
                                                       userInfo:userInfo];
                   
                   
                   failure(newError);
               }];
    }];
    
}

+ (void)patchOperation:(NSString*)path
          customHeader:(NSDictionary*)customHeader
            customBody:(NSString*)bodyString
               success:(void (^)(id responseHeader, id responseObject))success
               failure:(void (^)(id responseObject))failure{
    [[AuthenticationManager sharedInstance] checkAndRefreshToken:^(ADAuthenticationError *error) {
        if(error){
            failure(error);
            return;
        }
        NSMutableURLRequest *request = [[NSMutableURLRequest alloc] init];
        [request setURL:[NSURL URLWithString:path]];
        [request setHTTPMethod:@"PATCH"];
        
        [request addValue:[NSString stringWithFormat:@"Bearer %@", [[AuthenticationManager sharedInstance] accessToken]]
       forHTTPHeaderField:@"Authorization"];
        
        //set headers
        NSArray *customHeaderArray = [customHeader allKeys];
        for(NSString *key in customHeaderArray){
            [request addValue:[customHeader objectForKey:key] forHTTPHeaderField:key];
        }
        
        //create the body
        NSMutableData *postBody = [NSMutableData data];
        [postBody appendData:[bodyString dataUsingEncoding:NSUTF8StringEncoding]];
        
        //post
        [request setHTTPBody:postBody];
        
        AFHTTPRequestOperation *operation = [[AFHTTPRequestOperation alloc]
                                             initWithRequest:request];
        operation.responseSerializer = [AFJSONResponseSerializer serializer];
        
        [operation setCompletionBlockWithSuccess:^(AFHTTPRequestOperation *operation, id responseObject) {
            NSLog(@"status code = %ld", (long)operation.response.statusCode);
            
            success(operation.response.allHeaderFields, responseObject);
        } failure:^(AFHTTPRequestOperation *operation, NSError *error) {
            NSMutableDictionary *userInfo = [NSMutableDictionary dictionaryWithDictionary:error.userInfo];
            
            if(![operation.responseString isEqual:[NSNull null]])
                [userInfo setObject:operation.responseString forKey:@"responseString"];
            
            NSError *newError = [NSError errorWithDomain:error.domain
                                                    code:error.code
                                                userInfo:userInfo];
            
            
            failure(newError);
        }];
        [operation start];
    }];
}

+ (void)post:(NSString*)path
customHeader:(NSDictionary*)customHeader
  customBody:(NSString*)bodyString
     success:(void (^)(id responseHeader, id responseObject))success
     failure:(void (^)(id responseObject))failure{
    [[AuthenticationManager sharedInstance] checkAndRefreshToken:^(ADAuthenticationError *error) {
        if(error){
            failure(error);
            return;
        }
        NSMutableURLRequest *request = [[NSMutableURLRequest alloc] init];
        [request setURL:[NSURL URLWithString:path]];
        [request setHTTPMethod:@"POST"];
        
        [request addValue:[NSString stringWithFormat:@"Bearer %@", [[AuthenticationManager sharedInstance] accessToken]]
       forHTTPHeaderField:@"Authorization"];
        
        //set headers
        NSArray *customHeaderArray = [customHeader allKeys];
        for(NSString *key in customHeaderArray){
            [request addValue:[customHeader objectForKey:key] forHTTPHeaderField:key];
        }
        //create the body
        NSMutableData *postBody = [NSMutableData data];
        [postBody appendData:[bodyString dataUsingEncoding:NSUTF8StringEncoding]];
        
        //post
        [request setHTTPBody:postBody];
        
        AFHTTPRequestOperation *operation = [[AFHTTPRequestOperation alloc]
                                             initWithRequest:request];
        operation.responseSerializer = [AFJSONResponseSerializer serializer];
        
        [operation setCompletionBlockWithSuccess:^(AFHTTPRequestOperation *operation, id responseObject) {
            NSLog(@"status code = %ld", (long)operation.response.statusCode);
            
            success(operation.response.allHeaderFields, responseObject);
        } failure:^(AFHTTPRequestOperation *operation, NSError *error) {
            NSMutableDictionary *userInfo = [NSMutableDictionary dictionaryWithDictionary:error.userInfo];
            
            if(![operation.responseString isEqual:[NSNull null]])
                [userInfo setObject:operation.responseString forKey:@"responseString"];
            
            NSError *newError = [NSError errorWithDomain:error.domain
                                                    code:error.code
                                                userInfo:userInfo];
            
            
            failure(newError);
        }];
        [operation start];
    }];
}

+ (void)postWithMultipartForm:(NSString*)path
                  queryParams:(NSDictionary*)queryParams
             multiformObjects:(NSArray*)multiformObjects
                      success:(void (^)(id responseHeader, id responseObject))success
                      failure:(void (^)(id responseObject))failure{
    [[AuthenticationManager sharedInstance] checkAndRefreshToken:^(ADAuthenticationError *error) {
        if(error){
            failure(error);
            return;
        }
        AFHTTPRequestOperationManager *manager = [AFHTTPRequestOperationManager manager];
        
        [manager.requestSerializer setValue:[NSString stringWithFormat:@"Bearer %@", [[AuthenticationManager sharedInstance] accessToken]]
                         forHTTPHeaderField:@"Authorization"];
        
        
        [manager POST:path parameters:[self paramsRemove:self.keysToRemove from:queryParams] constructingBodyWithBlock:^(id<AFMultipartFormData> formData) {
            
            for (MultiformObject *obj in multiformObjects){
                [formData appendPartWithHeaders:obj.headers body:obj.body];
            }
        } success:^(AFHTTPRequestOperation *operation, id responseObject) {
            success(operation.response.allHeaderFields, responseObject);
            
        } failure:^(AFHTTPRequestOperation *operation, NSError *error) {
            NSMutableDictionary *userInfo = [NSMutableDictionary dictionaryWithDictionary:error.userInfo];
            
            if(![operation.responseString isEqual:[NSNull null]])
                [userInfo setObject:operation.responseString forKey:@"responseString"];
            
            NSError *newError = [NSError errorWithDomain:error.domain
                                                    code:error.code
                                                userInfo:userInfo];
            
            
            failure(newError);
        }];
    }];
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
