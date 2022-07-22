
# Sharepoint http data connector

Sharepoint http data connector is a library that can help us manage API calls to our sites on Sharepoint.
Provides a list of features that are useful for this integration.

## Overview

Sharepoint http data connector is an Open-Source project, remember that you can support it if you want.

During the documentation you will se that is simple the implementation of this library, and can help
you to much when you need to apply a custom integration with a site in Sharepoint.

Before to start, remember that there is some information that is very important to have to use the library.
The data that we need is the following:

- **Authentication information**:
    - AuthenticationUrl
    - TenantId
    - ClientId
    - ClientSecret
    - GrantType
    - Resource

- **Sharepoint instance information**:
    - SharepointSiteId
    - SharepointSiteName
    - SharepointInstanceUrl
    - SharepointSiteUrl
    - ServerRelativeUrl

Unfortunately, the current authentication type that support this library is ***client_credentials***, but
I foresee to incorporate different ways to authenticate in the future.

*If you don't know how to recover this information, don't worry!!*  
*You can find all the information that you need with the following link(s):*

- https://www.youtube.com/watch?v=uHwQzeyDc1A

## Documentation

We can implement the use of this library for different kind of projects, for example a sample demo or with 
ASP.NET Core.

### Sample implementation

To implement this library, one option is only making an instance to the class ***SharepointDataContext*** 
as shown below:

```c#
    SharepointDataContext sharepointContext = new SharepointDataContext(new SharepointContextConfiguration
    {
        // Authentication for SharePoint configuration.
        AuthenticationUrl = "https://accounts.accesscontrol.windows.net/{tenantid}/",
        TenantId = "00000000-0000-0000-0000-000000000000",
        ClientId = "00000000-0000-0000-0000-000000000000@00000000-0000-0000-0000-000000000000",
        ClientSecret = "{client_secret}",
        GrantType = "client_credentials",
        Resource = "00000003-0000-0ff1-ce00-000000000000/{your_organization}.sharepoint.com@{tenantid}",
        // Rest API to sharepoint site configuration.
        SharepointSiteId = "00000000-0000-0000-0000-000000000000",
        SharepointSiteName = "{site_name}",
        SharepointInstanceUrl = "https://{your_organization}.sharepoint.com/",
        SharepointSiteUrl = "https://{your_organization}.sharepoint.com/sites/{your_site}",
        ServerRelativeUrl = "/sites/{your_site}/" | "/sites/{your_site}/{library_folder_path}/" | "/"
    });
```

The class ***SharepointContextConfiguration*** contains the main information to connect correctly with an
instance of SharePoint.

### ASP.NET Core implementation

As the previous example, the implementation in ASP.NET Core project is similar, we only need to register
our service in our class ***program.cs*** as is shown in the following script.

```c#
    builder.Services.AddScoped<ISharepointDataContext>(s => new SharepointDataContext(new SharepointContextConfiguration
    {
        // Authentication for SharePoint configuration.
        AuthenticationUrl = builder.Configuration.GetSection("SharepointSite:Authentication:authenticationUrl").Value,
        TenantId = builder.Configuration.GetSection("SharepointSite:Authentication:tenantId").Value,
        ClientId = builder.Configuration.GetSection("SharepointSite:Authentication:clientId").Value,
        ClientSecret = builder.Configuration.GetSection("SharepointSite:Authentication:clientSecret").Value,
        GrantType = builder.Configuration.GetSection("SharepointSite:Authentication:grantType").Value,
        Resource = builder.Configuration.GetSection("SharepointSite:Authentication:resource").Value,
        // Rest API to SharePoint site configuration.
        SharepointSiteId = builder.Configuration.GetSection("SharepointSite:RestAPI:id").Value,
        SharepointSiteName = builder.Configuration.GetSection("SharepointSite:RestAPI:name").Value,
        SharepointInstanceUrl = builder.Configuration.GetSection("SharepointSite:RestAPI:resource").Value,
        SharepointSiteUrl = builder.Configuration.GetSection("SharepointSite:RestAPI:site").Value,
        ServerRelativeUrl = builder.Configuration.GetSection("SharepointSite:RestAPI:serverRelativeUrl").Value
    }));
```

***Note:*** *To have a better structure in our code, we can add this service as an extension method.*

In our ***Controller*** class we need to add the following code to make use of ***Dependency Injection*** with this library.

```c#
    public class ExampleController : ControllerBase
    {
        private readonly ISharepointDataContext _sharepointContext;

        public ExampleController(ISharepointDataContext sharepointContext) 
        {
            _sharepointContext = sharepointContext;
        }
    }
```

***Note:*** *Also, we can create our custom instance class of **SharepointDataContext** class to only make use of
the services that we need and limit the access to services in our Sharepoint site.*

### SharePoint http data connector services

In this section we will explain and show some examples of the use of this library.   

**Remember**   
You will see that our library, some functions require a **ServerRelativeUrl**, it is not more than the path or 
route, of our folder or file.
Here are some examples.  
- *my-folder/graduation*  
- *my-folder/photos*

If we don't set a value in the attribute ***ServerRelativeUrl*** of the library configuration when we do an instance of it,
it will take the default base address instead a custom relative path, this can help us if we only want to have an 
interaction with only one library documents in SharePoint.

```c#
    // Returns a true or false if the folder exists or not.
    bool response = await _sharepointContext.ExistsFolderAsync("my-folder/graduation");

    // Returns a true or false if the file exists or not.
    bool response = await _sharepointContext.ExistsFileAsync("my-folder/graduation", "graduationFile.pdf");

    // Returns a record type Sharepoint Recycle Resource by a unique identifier.
    var response = await _sharepointContext.GetRecycleBinResourceByIdAsync(new Guid("00000000-0000-0000-0000-000000000000"));

    // Returns the file content of a document in SharePoint.
    byte[]? response = await _sharepointContext.DownloadFileAsync("my-folder/graduation", "graduationFile.pdf");
    
    // Returns true or false if the resource was deleted or not (Support file name in relative url).
    bool response = await _sharepointContext.DeleteResourceAsync("my-folder/graduation");
    bool response = await _sharepointContext.DeleteResourceAsync("my-folder/graduation/graduationFile.pdf");

    // Returns true or false if a file was deleted or not.
    bool response = await _sharepointContext.DeleteFileAsync("my-folder/graduation", "graduationFile.pdf");

    // Returns a record type Sharepoint Folder when it's created.
    // The difference between these two functions is that the first one, create a folder in the main server relative
    // url that was defined in the Configuration of the library, the second one, let us create a folder in a specific route. 
    var response = await _sharepointContext.CreateFolderAsync("my-new-Folder");
    var response = await _sharepointContext.CreateFolderAsync("my-folder/graduation", "my-new-folder");
    
    // Retuns a record type Sharepoint File when one is uploaded.
    var response = await _sharepointContext.UploadFileAsync("my-folder/graduation", "graduationPhoto.jpg", byte[] content);

    // Returns an unique identifier when a you want to move a resource to the Recycle bin (Support file name in relative url).
    var recordId = await _sharepointContext.DeleteResourceToRecycleBinByIdAsync("my-folder/graduation/graduationPhoto.jpg");

    // Returns true or false if the resource was restored correctly from Recycle bin.
    bool response = await _sharepointContext.RestoreRecycleBinResourceByIdAsync(new Guid("00000000-0000-0000-0000-000000000000"));
```

For future updates in the library, we plan to cover many other functions that can be useful in our interaction with
a SharePoint Site.

## Repository

- [Sharepoint.Http.Data.Connector](https://github.com/orion-it-solutions/sharepoint-http-data-connector/tree/develop)

## Authors

- [@emmanueltoledo](https://github.com/emmanuel-toledo)


