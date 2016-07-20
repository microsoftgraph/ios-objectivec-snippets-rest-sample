# Office 365 iOS Microsoft Graph Snippets



## Introduction

Microsoft Graph (previously called Office 365 unified API) exposes multiple APIs from Microsoft cloud services through a single REST API endpoint. This repository shows you how to access multiple resources, including Microsoft Azure Active Directory (AD) and the Office 365 APIs, by making HTTP requests to the Microsoft Graph API in an iOS application. 

> Note: Try out the [Get started with Office 365 APIs](http://dev.office.com/getting-started/office365apis?platform=option-ios#setup) page which simplifies registration so you can get this sample running faster.
 
## Prerequisites
* [Xcode](https://developer.apple.com/xcode/downloads/) from Apple
* Installation of [CocoaPods](https://guides.cocoapods.org/using/using-cocoapods.html)  as a dependency manager.
* An Office 365 account. You can sign up for [an Office 365 Developer subscription](https://aka.ms/devprogramsignup) that includes the resources that you need to start building Office 365 apps.

     > Note: If you already have a subscription, the previous link sends you to a page with the message *Sorry, you can’t add that to your current account*. In that case, use an account from your current Office 365 subscription.
* A Microsoft Azure tenant to register your application. Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

     > Important: You will also need to ensure your Azure subscription is bound to your Office 365 tenant. To do this see the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). The section **Adding a new directory** will explain how to do this. You can also see [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) and the section **Associate your Office 365 account with Azure AD to create and manage apps** for more information.
      
* A client id and redirect uri values of an application registered in Azure. This sample must be registered and granted specific permissions for **Microsoft Graph**. To learn how to create this registration, see [Register your native app with the Azure Management Portal](https://msdn.microsoft.com/office/office365/howto/add-common-consent-manually). Next, to assign the correct permissions to the registration, see [grant proper permissions](https://github.com/OfficeDev/O365-iOS-Microsoft-Graph-Snippets/wiki/Grant-permissions-to-the-Snippets-application-in-Azure) in the repository wiki. 


## Running this sample in Xcode

1. Clone this repository
2. Use CocoaPods to import the necessary dependencies:
        
	     pod 'AFNetworking', '2.6.0'
	     pod 'ADALiOS', '1.2.4'

	This sample app already contains a podfile that will get the necessary components(pods) into  the project. Simply navigate to the project from **Terminal** and run 
        
        pod install
        
   	For more information, see **Using CocoaPods** in [Additional Resources](#AdditionalResources)
  
3. Open **O365-iOS-Microsoft-Graph-Snippets.xcworkspace**
4. Open **ConnectViewController.m** under the **Controllers** folder in the project. Ensure you have registered your app in Microsoft Azure, configured the proper permissions as specified, and obtained the **ClientID** and **RedirectUri** values from the registration. Again, this procedure is detailed under the fifth bullet point in the prerequisites section: **A client id and redirect uri values of an application registered in Azure.** You'll see that the **ClientID** and **RedirectUri** values can be added to the top of the file. Supply the necessary values here:

        // You will set your application's clientId and redirect URI. You get
        // these when you register your application in Microsoft Azure.
        NSString * const kRedirectUri = @"ENTER_REDIRECT_URI_HERE";
        NSString * const kClientId    = @"ENTER_CLIENT_ID_HERE";
        NSString * const kResourceId  = @"https://graph.microsoft.com";
        NSString * const kAuthority   = @"https://login.microsoftonline.com/common";
    
    > Note: If you have don't have CLIENT_ID and REDIRECT_URI values, see [Register your native app with the Azure Management Portal](https://msdn.microsoft.com/office/office365/howto/add-common-consent-manually) and take note of the CLIENT\_ID and REDIRECT_URI.

5.Run the sample

To learn more about the sample, visit our [understanding the code](https://github.com/OfficeDev/O365-iOS-Microsoft-Graph-Snippets/wiki/Understanding-the-Snippets-sample-code) wiki page.


## Questions and comments

We'd love to get your feedback about the Office 365 iOS Snippets project. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/O365-iOS-Microsoft-Graph-Snippets/issues) section of this repository.

Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [Office365] and [MicrosoftGraph].

## Contributing
You will need to sign a [Contributor License Agreement](https://cla.microsoft.com/) before submitting your pull request. To complete the Contributor License Agreement (CLA), you will need to submit a request via the form and then electronically sign the CLA when you receive the email containing the link to the document. 

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Additional resources

* [Office Dev Center](http://dev.office.com/)
* [Microsoft Graph overview page](https://graph.microsoft.io)
* [Using CocoaPods](https://guides.cocoapods.org/using/using-cocoapods.html)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
