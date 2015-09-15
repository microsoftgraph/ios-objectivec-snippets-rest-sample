/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import "ConnectViewController.h"
#import "AuthenticationManager.h"
#import <ADAuthenticationResult.h>
#import <ADAuthenticationError.h>
#import "DetailTableViewController.h"


// You will set your application's clientId and redirect URI.
NSString * const kRedirectUri = @"ENTER_REDIRECT_URI_HERE";
NSString * const kClientId    = @"ENTER_CLIENT_ID_HERE";
NSString * const kAuthority   = @"https://login.microsoftonline.com/common";

@interface ConnectViewController () <UISplitViewControllerDelegate>

@property (weak, nonatomic) IBOutlet UIActivityIndicatorView *activityIndicator;
@property (weak, nonatomic) IBOutlet UIButton *connectButton;

@end

@implementation ConnectViewController

- (void)viewDidLoad {
    [super viewDidLoad];
}


- (void) viewDidAppear:(BOOL)animated{
    [super viewDidAppear:animated];
   
}

- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

- (IBAction)authenticateWithAzureTapped:(id)sender {
    [self showLoadingUI:YES];
    AuthenticationManager *authManager = [AuthenticationManager sharedInstance];
    
    [authManager initWithAuthority:kAuthority
                          clientId:kClientId
                       redirectURI:kRedirectUri
                        resourceID:@"https://graph.microsoft.com"
                        completion:^(ADAuthenticationError *error) {
                            if(error){
                                [self showLoadingUI:NO];
                                [self handleADAuthenticationError:error];
                            }
                            else{
                                [authManager acquireAuthTokenCompletion:^(ADAuthenticationError *acquireTokenError) {
                                    if(acquireTokenError){
                                        [self showLoadingUI:NO];
                                        [self handleADAuthenticationError:acquireTokenError];
                                    }
                                    else{
                                        [self showLoadingUI:NO];
                                        [[NSOperationQueue mainQueue] addOperationWithBlock:^{
                                            [self performSegueWithIdentifier:@"showSplitView" sender:nil];
                                            
                                        }];
                                    }
                                }];
                            }
                        }];
}

#pragma mark - helper
- (void)showLoadingUI:(BOOL)loading{
    if(loading){
        [self.activityIndicator startAnimating];
        self.connectButton.enabled = NO;
    }
    else{
        [self.activityIndicator stopAnimating];
        self.connectButton.enabled = YES;
    }
}

- (void)handleADAuthenticationError:(ADAuthenticationError *)error{
    NSLog(@"Error\nProtocol Code %@\nDescription %@", error.protocolCode, error.description);
    UIAlertController *alert = [UIAlertController alertControllerWithTitle:@"Error"
                                                                   message:@"Please see the log for more details"
                                                            preferredStyle:UIAlertControllerStyleAlert];
    [alert addAction:[UIAlertAction actionWithTitle:@"Close"
                                              style:UIAlertActionStyleDefault handler:^(UIAlertAction *action) {
                                                  ;
                                              }]];
    [self presentViewController:alert animated:YES completion:^{
        ;
    }];
}

#pragma mark - Navigation

// In a storyboard-based application, you will often want to do a little preparation before navigation
- (void)prepareForSegue:(UIStoryboardSegue *)segue sender:(id)sender {
    // Get the new view controller using [segue destinationViewController].
    // Pass the selected object to the new view controller.
    UISplitViewController *splitViewController = segue.destinationViewController;
    splitViewController.delegate = self;
}


- (BOOL)splitViewController:(UISplitViewController *)splitViewController collapseSecondaryViewController:(UIViewController *)secondaryViewController ontoPrimaryViewController:(UIViewController *)primaryViewController {
    if ([secondaryViewController isKindOfClass:[UINavigationController class]] && [[(UINavigationController *)secondaryViewController topViewController] isKindOfClass:[DetailTableViewController class]] && ([(DetailTableViewController *)[(UINavigationController *)secondaryViewController topViewController] operation] == nil)) {
        // Return YES to indicate that we have handled the collapse by doing nothing; the secondary controller will be discarded.
        return YES;
    } else {
        return NO;
    }
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
