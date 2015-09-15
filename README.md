# O365-iOS-Unified-API-Snippets-Preview



## Introduction

The Office 365 unified API exposes multiple APIs from Microsoft cloud services through a single REST API endpoint. This repository shows you how to access multiple resources, including Microsoft Azure Active Directory (AD) and the Office 365 APIs, by making HTTP requests to the Office 365 unified API in an iOS application. 


**Note: If possible, please use this sample with a "non-work" or test account in Office 365. With the current version of the project, it does not clean up the created objects in your mailbox, calendar, contacts, and objects created from additional operations. At this time you'll have to manually remove these artifacts, for example, sample mails, contacts, and calendar events.**  

## Prerequisites
* [Xcode](https://developer.apple.com/xcode/downloads/) from Apple
* Installation of [CocoaPods](https://guides.cocoapods.org/using/using-cocoapods.html)  as a dependency manager.
* An Office 365 account. You can sign up for [an Office 365 Developer subscription](https://portal.office.com/Signup/Signup.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK&ali=1#0) that includes the resources that you need to start building Office 365 apps.

     > Note: If you already have a subscription, the previous link sends you to a page with the message *Sorry, you can’t add that to your current account*. In that case, use an account from your current Office 365 subscription.
* A Microsoft Azure tenant to register your application. Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

     > Important: You will also need to ensure your Azure subscription is bound to your Office 365 tenant. To do this see the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). The section **Adding a new directory** will explain how to do this. You can also see [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) and the section **Associate your Office 365 account with Azure AD to create and manage apps** for more information.
      
* A client id and redirect uri values of an application registered in Azure. This sample must be registered and granted specific permissions for the **Office 365 unified API (preview)**. To learn how to create this registration, see [Add a native client application in Azure](https://msdn.microsoft.com/library/azure/dn132599.aspx#BKMK_Adding). Next, to assign the correct permissions to the registration, see [grant proper permissions](https://github.com/OfficeDev/O365-iOS-Unified-API-Snippets/wiki/Grant-permissions-to-the-Snippets-application-in-Azure) in the repository wiki. 


## Opening the sample using Xcode

1. Clone this repository
2. Use CocoaPods to import the Active Directory Authentication Library (ADAL) iOS dependency:
        
	     pod 'ADALiOS', '~> 1.2.4'

 This sample app already contains a podfile that will get the ADAL components(pods) into  the project. Simply navigate to the project from **Terminal** and run 
        
        pod install
        
   For more information, see **Using CocoaPods** in [Additional Resources](#AdditionalResources)
  
3. Open **O365-iOS-Unified-API-Snippets.xcworkspace**
4. Open **ConnectViewController.m**. Ensure you have completed the third bullet point in the prerequisites section, including registering your app in Microsoft Azure, configuring the proper permissions as specified, and obtaining the **ClientID** and **RedirectUri** values from the registration. You'll see that the **ClientID** and **RedirectUri** values can be added to the top of the file. Supply the necessary values here:

        // You will set your application's clientId and redirect URI. You get
        // these when you register your application in Microsoft Azure.
        NSString * const kRedirectUri = @"ENTER_REDIRECT_URI_HERE";
        NSString * const kClientId    = @"ENTER_CLIENT_ID_HERE";
        NSString * const kAuthority   = @"https://login.microsoftonline.com/common";
    
    > Note: If you have don't have CLIENT_ID and REDIRECT_URI values, [add a native client application in Azure](https://msdn.microsoft.com/library/azure/dn132599.aspx#BKMK_Adding) and take note of the CLIENT\_ID and REDIRECT_URI.

Run the sample once above steps are complete.

To learn more about the sample, visit our [understanding the code](https://github.com/OfficeDev/O365-iOS-Unified-API-Snippets/wiki) wiki page.


## Questions and comments

We'd love to get your feedback about the Office 365 iOS Snippets project. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/O365-iOS-Unified-API-Snippets/issues) section of this repository.

Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [Office365] and [API].


## Additional resources

* [Office 365 Connect Sample for iOS Using Unified API (Preview)](O365-iOS-Unified-API-Connect)
* [Office 365 unified API overview (preview)](https://msdn.microsoft.com/en-us/office/office365/howto/office-365-unified-api-overview)
* [Microsoft Azure Active Directory Authentication Library (ADAL) for iOS and OSX](https://github.com/AzureAD/azure-activedirectory-library-for-objc)

