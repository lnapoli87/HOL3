Module XX: *Manage Lists in a O365 tenant with iOS*
==========================

##Overview

The lab lets students use an AzureAD account to manage files in a O365 Sharepoint tenant with an iOS app.

##Objectives

- Learn how to create a client for O365 to list files and download to the local storage to then show it in a preview page.

##Prerequisites

- Apple Macintosh environment
- XCode 6 (from the AppStore - compatible with iOS8 sdk)
- XCode developer tools (it will install git integration from XCode and the terminal)
- You must have a Windows Azure subscription to complete this lab.
- You must have completed Module 04 and linked your Azure subscription with your O365 tenant.

##Exercises

The hands-on lab includes the following exercises:

- [Add O365 iOS files sdk library to the project](#exercise1)
- [Create a Client class for all operations](#exercise2)
- [Connect actions in the view to ProjectClient class](#exercise3)

<a name="exercise1"></a>
##Exercise 1: Add O365 iOS files sdk library to a project
In this exercise you will use an existing application with the AzureAD authentication included, to add the O365 files sdk library in the project
and create a client class with empty methods in it to handle the requests to the Sharepoint tenant.

###Task 1 - Open the Project
01. Download the starting point App:

    ```
    git clone 
    ```

02. Open the **.xcodeproj** file in the O365-Files-App

03. Find and Open the **ViewController.m** class under **O365-lists-app/controllers/login/**

04. Fill the AzureAD account settings in the **viewDidLoad** method
    
    ![](img/fig.01.png)

03. Build and Run the project in an iOS Simulator to check the views

    ```
    Application:
    You will se a login page with buttons to access the application and to clear credentials.
    Once authenticated, a File list will appear with one fake entry. Also there is a File 
    Details screen (selecting a row in the table) with the name, last modified and created dates.
    Finally, there is an action button to download the File.

    Environment:
    To access the files, in the O365 Sharepoint tenant there is a Default space to store documents
    called "Shared Documents". We will use the o365-files-sdk to access these files, download them,
    and show a preview in the iOS application.
    ```

    ![](img/fig.02.png)

###Task 2 - Importing the library
01. Download a copy of the library using the terminal:

    ```
    git clone 
    ```

02. Open the downloaded folder and copy **office365-files-sdk** folder under **Sdk-ObjectiveC**. Paste it in a lib folder inside our project path.

    ![](img/fig.03.png)

03. Drag the **office365-files-sdk.xcodeproj** file into XCode under our application project.
    
    ![](img/fig.04.png)

04. Repeat steps 02 and 03 with **office365-base-sdk**

05. Go to project settings selecting the first file from the files explorer. Then click on **Build Phases** and add an entry in the **Target Dependencies** section.

    ![](img/fig.05.png)

06. Select the **office365-files-sdk** and **office365-base-sdk** library dependencies.

    ![](img/fig.06.png)

07. Under **Link Binary with Libraries** add an entry pointing to **office365-base-sdk.a** and **office365-list-sdk.a** files

    ![](img/fig.07.png)

09. Now delete **ADALiOS.xcodeproj** from the project and select **Remove Reference** 
    
    ```
    This step avoids conflicts because office365-base-sdk already has ADALiOS 
    and is not necesary to have the library added twice
    ```    

    ![](img/fig.08.png)

08. Build and Run the application to check everything is ok.

    ![](img/fig.09.png)

<a name="exercise2"></a>
##Exercise 2: Create a Client class for all operations
In this exercise you will create a client class for operations related to Files. This class will connect to the **office365-files-sdk**, and will be subclass of **FileClient**.

###Task 1 - Create a client class to connect to the o365-files-sdk

01. On the XCode files explorer, make a right click in the group **Helpers** and select **New File**. You will see the **New File wizard**. Click on the **iOS** section, select **Cocoa Touch Class** and click **Next**.

    ![](img/fig.10.png)

03. In this section, configure the new class giving it a name (**CustomFileClient**), and make it a subclass of **ListClient**. Make sure that the language dropdown is set with **Objective-C** because our o365-lists library is written in that programming language. Finally click on **Next**.

    ![](img/fig.11.png)    

04. Now we are going to select where the new class sources files (.h and .m) will be stored. In this case we can click on **Create** directly. This will create a **.h** and **.m** files for our new class.

    ![](img/fig.12.png)

05. Build the Project and you will see one error. To fix it, change the import sentence On **CustomFileClient.h**.

    From :
    ```
    #import "FileClient.h"
    ```

    To:
    ```
    #import <office365-files-sdk/FileClient.h>
    ```

08. Re-build the project and check everything is ok.



###Task 2 - Add CustomFileClient methods

01. Open the **CustomFileClient.h** class and then add the following between **@interface** and **@end**

    ```
    - (NSURLSessionDataTask *)getFiles:(NSString *)folder callback :(void (^)(NSMutableArray *files, NSError *))callback;
    - (NSURLSessionDataTask *)download:(NSString *)fileName callback :(void (^)(NSData *data, NSError *error))callback;
    +(FileClient*)getClient:(NSString *) token;
    ```

02. Add the body of each method in the **CustomFileClient.m** file.

    Get Files
    ```
    const NSString *apiUrl = @"/_api/files";

- (NSURLSessionDataTask *)getFiles:(NSString *)folder callback :(void (^)(NSMutableArray *files, NSError *))callback{
    
    NSString *url;
    
    if(folder == nil){
        url = [NSString stringWithFormat:@"%@%@", self.Url , apiUrl];
    }
    else{
        url = [NSString stringWithFormat:@"%@%@", self.Url , apiUrl, [folder urlencode]];
    }
    
    HttpConnection *connection = [[HttpConnection alloc] initWithCredentials:self.Credential url:url];
    
    NSString *method = (NSString*)[[Constants alloc] init].Method_Get;
    
    return [connection execute:method callback:^(NSData  *data, NSURLResponse *reponse, NSError *error) {
        NSMutableArray *array = [NSMutableArray array];
        
        if(error == nil){
            array = [self parseData : data];
        }
        
        callback(array, error);
    }];
}
    ```

    Download File
    ```
- (NSURLSessionDataTask *)download:(NSString *)fileName callback :(void (^)(NSData *data, NSError *error))callback{
    
    NSString *url = [NSString stringWithFormat:@"%@%@('%@')/download", self.Url , apiUrl, [fileName stringByAddingPercentEscapesUsingEncoding:NSUTF8StringEncoding]];

    HttpConnection *connection = [[HttpConnection alloc] initWithCredentials:self.Credential url:url ];

    NSString *method = (NSString*)[[Constants alloc] init].Method_Get;

    return [connection execute:method callback:^(NSData  *data, NSURLResponse *reponse, NSError *error) {
        callback(data, error);
    }];
}
    ```

    
03. Add the **getClient** class method

    ```
    +(CustomFileClient*)getClient:(NSString *) token{
    OAuthentication* authentication = [OAuthentication alloc];
    [authentication setToken:token];
    
    return [[CustomFileClient alloc] initWithUrl:@"https://xxx.xxx/xxx"
                               credentials: authentication];
    }
    ```

    ```
    Make sure to change https://xxx.xxx/xxx with the Resource url in the 
    initWithUrl:credentials: method.
    ```

04. Add the following import sentences:

    ```
    #import "office365-base-sdk/NSString+NSStringExtensions.h"
    #import "office365-base-sdk/HttpConnection.h"
    #import "office365-base-sdk/Constants.h"
    #import "office365-base-sdk/OAuthentication.h"
    ```

05. Build the project and check everything is ok.


<a name="exercise3"></a>
##Exercise 3: Connect actions in the view to CustomFileClient class
In this exercise you will navigate in every controller class of the project, in order to connect each action (from buttons, lists and events) with one CustomFileClient operation.

```
The Application has every event wired up with their respective controller classes. 
We need to connect this event methods to our CustomFileClient class 
in order to have access to the o365-files-sdk.
```

###Task1 - Wiring up FileListView

01. Open **FileListViewController.h** class header and add a property to store the files.

    ```
    @property NSMutableArray *files;
    ```

    Also add an instance variable in the **FileListViewController.m** to hold the current selection
    ```
    FileEntity* currentEntity;
    ```


02. Open **FileListViewController.m** class implementation and the **loadData** method:

    ```
    -(void) loadData{
    //Create and add a spinner
    double x = ((self.navigationController.view.frame.size.width) - 20)/ 2;
    double y = ((self.navigationController.view.frame.size.height) - 150)/ 2;
    UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(x, y, 20, 20)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    [spinner startAnimating];
    
    CustomFileClient *client = [CustomFileClient getClient:self.token];
    NSURLSessionDataTask *task = [client getFiles:@"" callback:^(NSMutableArray *files, NSError *error) {
        self.files = files;
        dispatch_async(dispatch_get_main_queue(), ^{
            [self.tableView reloadData];
            [spinner stopAnimating];
        });
    }];
    [task resume];
}
    ```

    Now call it from the **viewWillAppear** method. Also add the initialization for **currentEntity** and **files**
    ```
    - (void)viewWillAppear:(BOOL)animated{
    [self loadData];
    currentEntity = nil;
    self.files = [[NSMutableArray alloc] init];
    }
    ```

03. Add the table methods:

```
- (NSInteger)tableView:(UITableView *)tableView numberOfRowsInSection:(NSInteger)section{
    return self.files.count;
}

- (UITableViewCell *)tableView:(UITableView *)tableView cellForRowAtIndexPath:(NSIndexPath *)indexPath{
    NSString* identifier = @"fileListCell";
    FileListCellTableViewCell *cell =[tableView dequeueReusableCellWithIdentifier: identifier ];
    
    FileEntity *file = [self.files objectAtIndex:indexPath.row];
    
    cell.fileName.text = file.Name;
    cell.lastModified.text = [NSString stringWithFormat:@"Last modified on %@", [file.TimeLastModified substringToIndex:10]];
    
    return cell;
}

- (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath{
    currentEntity= [self.files objectAtIndex:indexPath.row];
    
    [self performSegueWithIdentifier:@"detail" sender:self];
}
```

04. Add the navigation methods

```
- (BOOL)shouldPerformSegueWithIdentifier:(NSString *)identifier sender:(id)sender{
    return ([identifier isEqualToString:@"detail"] && currentEntity);
}

-(void) prepareForSegue:(UIStoryboardSegue *)segue sender:(id)sender{
    if([segue.identifier isEqualToString:@"detail"]){
        FileDetailsViewController *ctrl = (FileDetailsViewController *)segue.destinationViewController;
        //ctrl.token = self.token;
        //ctrl.file = currentEntity;
    }
}
```

05. Add the needed import sentences:

```
#import "FileListCellTableViewCell.h"
#import "office365-files-sdk/FileClient.h"
#import "office365-base-sdk/OAuthentication.h"
#import "office365-files-sdk/FileEntity.h"
#import "CustomFileClient.h"
#import "FileDetailsViewController.h"
```

06. Build and Run the application. Check everything is ok. Now you will be able to se the Files list from the O365 Sharepoint tenant

    ![](img/fig.13.png)





###Task2 - Wiring up CreateProjectView

01. Open **CreateViewController.m** and add the body to the **createProject** method

    ```
    -(void)createProject{
    if(![self.FileNameTxt.text isEqualToString:@""]){
        UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
        spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
        [self.view addSubview:spinner];
        spinner.hidesWhenStopped = YES;
        
        [spinner startAnimating];
        
        ProjectClient* client = [ProjectClient getClient:self.token];
        
        ListItem* newProject = [[ListItem alloc] init];
        
        NSDictionary* dic = [NSDictionary dictionaryWithObjects:@[@"Title",self.FileNameTxt.text] forKeys:@[@"_metadata",@"Title"]];
        [newProject initWithDictionary:dic];
        
        NSURLSessionTask* task = [client addProject:newProject callback:^(BOOL success, NSError *error) {
            if(error == nil){
                dispatch_async(dispatch_get_main_queue(), ^{
                    [spinner stopAnimating];
                    [self.navigationController popViewControllerAnimated:YES];
                });
            }else{
                NSString *errorMessage = [@"Add Project failed. Reason: " stringByAppendingString: error.description];
                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:errorMessage delegate:self cancelButtonTitle:@"Retry" otherButtonTitles:@"Cancel", nil];
                [alert show];
            }
        }];
        [task resume];
    }else{
        dispatch_async(dispatch_get_main_queue(), ^{
            UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:@"Complete all fields" delegate:self cancelButtonTitle:@"Ok" otherButtonTitles:nil, nil];
            [alert show];
        });
    }
}
    ```

02. Add the import sentence to the **ProjectClient** class

    ```
    #import "ProjectClient.h"
    ```

03. Build and Run the app, and check everything is ok. Now you can create a new project with the plus button in the left corner of the main screen

    ![](img/fig.18.png)


###Task3 - Wiring up ProjectDetailsView

01. Open **ProjectDetailsViewController.h** and add the following variables

    ```
    @property ListItem* project;
    @property ListItem* selectedReference;
    ```

    And add the import sentence
    ```
    #import "office365-lists-sdk/ListItem.h"
    ```

02. Set the value when the user selects a project in the list. On **ProjectTableViewController.m**

    Uncomment this line in the **prepareForSegue:sender:** method
    ```
    //controller.project = currentEntity;
    ```

03. Back to the **ProjectDetailsViewController.m** Set the fields and screen title text on the **viewDidLoad** method

    ```
    -(void)viewDidLoad{
    self.projectName.text = self.project.getTitle;
    self.navigationItem.title = self.project.getTitle;
    self.navigationItem.rightBarButtonItem.title = @"Done";
    self.selectedReference = false;
    self.projectNameField.hidden = true;
    
    
    [self loadData];
    }
    ```

04. Load the references

    Load data method
    ```
    -(void)loadData{
    //Create and add a spinner
    UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    [spinner startAnimating];
    
    ProjectClient* client = [ProjectClient getClient:self.token];
    
    NSURLSessionTask* task = [client getList:@"Research References" callback:^(ListEntity *list, NSError *error) {
        
        //If list doesn't exists, create one with name Research References
        if(list){
            dispatch_async(dispatch_get_main_queue(), ^{
                [self getReferences:spinner];
            });
        }else{
            dispatch_async(dispatch_get_main_queue(), ^{
                [self createReferencesList:spinner];
            });
        }
        
    }];
    [task resume];
    
}
    ```

    Get References Method
    ```
        -(void)getReferences:(UIActivityIndicatorView *) spinner{
    ProjectClient* client = [ProjectClient getClient:self.token];
    
    NSURLSessionTask* listReferencesTask = [client getReferencesByProjectId:self.project.Id callback:^(NSMutableArray *listItems, NSError *error) {
            dispatch_async(dispatch_get_main_queue(), ^{
                self.references = [listItems copy];
                [self.refencesTable reloadData];
                [spinner stopAnimating];
            });
        
        }];

    [listReferencesTask resume];
}
    ```

    Create References Lists if not exists
    ```
    -(void)createReferencesList:(UIActivityIndicatorView *) spinner{
    ProjectClient* client = [ProjectClient getClient:self.token];
    
    ListEntity* newList = [[ListEntity alloc ] init];
    [newList setTitle:@"Research References"];
    
    NSURLSessionTask* createProjectListTask = [client createList:newList :^(ListEntity *list, NSError *error) {
        [spinner stopAnimating];
    }];
    [createProjectListTask resume];
}
    ```

05. Fill the table cells

    ```
    - (UITableViewCell *)tableView:(UITableView *)tableView cellForRowAtIndexPath:(NSIndexPath *)indexPath
{
    NSString* identifier = @"referencesListCell";
    ReferencesTableViewCell *cell =[tableView dequeueReusableCellWithIdentifier: identifier ];
    
    ListItem *item = [self.references objectAtIndex:indexPath.row];
    NSDictionary *dic =[item getData:@"URL"];
    cell.titleField.text = [dic valueForKey:@"Description"];
    cell.urlField.text = [dic valueForKey:@"Url"];
    
    return cell;
}
    ```

06. Get the references count
    
    ```
    - (NSInteger)tableView:(UITableView *)tableView numberOfRowsInSection:(NSInteger)section
{
    return [self.references count];
}
    ```

07. Row selection

    ```
    - (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath
{
    self.selectedReference= [self.references objectAtIndex:indexPath.row];    
    [self performSegueWithIdentifier:@"referenceDetail" sender:self];
}
    ```

08. Forward navigation

    ```
    - (void)prepareForSegue:(UIStoryboardSegue *)segue sender:(id)sender
{
    if([segue.identifier isEqualToString:@"createReference"]){
        CreateReferenceViewController *controller = (CreateReferenceViewController *)segue.destinationViewController;
        controller.project = self.project;
        controller.token = self.token;
    }else if([segue.identifier isEqualToString:@"referenceDetail"]){
        ReferenceDetailsViewController *controller = (ReferenceDetailsViewController *)segue.destinationViewController;
        controller.selectedReference = self.selectedReference;
        controller.token = self.token;
    }else if([segue.identifier isEqualToString:@"editProject"]){
        EditProjectViewController *controller = (EditProjectViewController *)segue.destinationViewController;
        controller.project = self.project;
        controller.token = self.token;
    }
    self.selectedReference = false;
}
    ```

    ```
    - (BOOL)shouldPerformSegueWithIdentifier:(NSString *)identifier sender:(id)sender{
    return ([identifier isEqualToString:@"referenceDetail"] && self.selectedReference) || [identifier isEqualToString:@"createReference"] || [identifier isEqualToString:@"editProject"];
}
    ```

09. Add the import sentence to the **ProjectClient** class

    ```
    #import "ProjectClient.h"
    ```

10. Build and Run the app, and check everything is ok. Now you can see the references from a project

    ![](img/fig.19.png)

###Task4 - Wiring up EditProjectView

01. Adding a variable for the selected project 

    First, add a variable **project** in the **EditProjectViewController.h**
    ```
    @property ListItem* project;
    ```

    And the import sentence
    ```
    #import "office365-lists-sdk/ListItem.h"
    ```

02. On the **ProjectDetailsViewController.m**, uncomment this line in the **prepareForSegue:sender:** method

    ```
    //controller.project = self.project;
    ```

03. Back to **EditProjectViewController.m**. Add the body for **updateProject**

    ```
    -(void)updateProject{
    if(![self.ProjectNameTxt.text isEqualToString:@""]){
        UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
        spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
        [self.view addSubview:spinner];
        spinner.hidesWhenStopped = YES;
        
        [spinner startAnimating];
        
        ListItem* editedProject = [[ListItem alloc] init];
        
        NSDictionary* dic = [NSDictionary dictionaryWithObjects:@[@"Title",self.ProjectNameTxt.text, self.project.Id] forKeys:@[@"_metadata",@"Title",@"Id"]];
        [editedProject initWithDictionary:dic];
        
        ProjectClient* client = [ProjectClient getClient:self.token];
        
        NSURLSessionTask* task = [client updateProject:editedProject callback:^(BOOL result, NSError *error) {
            if(error == nil){
                dispatch_async(dispatch_get_main_queue(), ^{
                    [spinner stopAnimating];
                    ProjectTableViewController *View = [self.navigationController.viewControllers objectAtIndex:self.navigationController.viewControllers.count-3];
                    [self.navigationController popToViewController:View animated:YES];
                });
            }else{
                NSString *errorMessage = [@"Update Project failed. Reason: " stringByAppendingString: error.description];
                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:errorMessage delegate:self cancelButtonTitle:@"Retry" otherButtonTitles:@"Cancel", nil];
                [alert show];
            }
        }];
        [task resume];
        
    }else{
        dispatch_async(dispatch_get_main_queue(), ^{
            UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:@"Complete all fields" delegate:self cancelButtonTitle:@"Ok" otherButtonTitles:nil, nil];
            [alert show];
        });
    }
}
    ```

04. Do the same for **deleteProject**

    ```
    -(void)deleteProject{
    UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    
    [spinner startAnimating];
    
    ProjectClient* client = [ProjectClient getClient:self.token];

    NSURLSessionTask* task = [client deleteListItem:@"Research Projects" itemId:self.project.Id callback:^(BOOL result, NSError *error) {
        if(error == nil){
            dispatch_async(dispatch_get_main_queue(), ^{
                [spinner stopAnimating];
                
                ProjectTableViewController *View = [self.navigationController.viewControllers objectAtIndex:self.navigationController.viewControllers.count-3];
                [self.navigationController popToViewController:View animated:YES];
            });
        }else{
            NSString *errorMessage = [@"Delete Project failed. Reason: " stringByAppendingString: error.description];
            UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:errorMessage delegate:self cancelButtonTitle:@"Retry" otherButtonTitles:@"Cancel", nil];
            [alert show];
        }
    }];
    
    [task resume];
}
    ```

05. Set the **viewDidLoad** initialization

    ```
    -(void)viewDidLoad{
    self.projectName.text = self.project.getTitle;
    self.navigationItem.title = self.project.getTitle;
    self.navigationItem.rightBarButtonItem.title = @"Done";
    self.selectedReference = false;
    self.projectNameField.hidden = true;
    
    
    [self loadData];
    }
    ```

06. Add the import sentence to the **ProjectClient** class

    ```
    #import "ProjectClient.h"
    ```

07. Build and Run the app, and check everything is ok. Now you can edit a project

    ![](img/fig.20.png)


###Task5 - Wiring up CreateReferenceView

01. On the **CreateReferenceViewController.m** add the body for the **createReference** method

    ```
    -(void)createReference{
    if((![self.referenceUrlTxt.text isEqualToString:@""]) && (![self.referenceDescription.text isEqualToString:@""]) && (![self.referenceTitle.text isEqualToString:@""])){
        UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
        spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
        [self.view addSubview:spinner];
        spinner.hidesWhenStopped = YES;
        
        [spinner startAnimating];
        
        ProjectClient* client = [self getClient];
        
        NSString* obj = [NSString stringWithFormat:@"{'Url':'%@', 'Description':'%@'}", self.referenceUrlTxt.text, self.referenceTitle.text];
        NSDictionary* dic = [NSDictionary dictionaryWithObjects:@[obj, self.referenceDescription.text, [NSString stringWithFormat:@"%@", self.project.Id]] forKeys:@[@"URL", @"Comments", @"Project"]];
        
        ListItem* newReference = [[ListItem alloc] initWithDictionary:dic];
        
        NSURLSessionTask* task = [client addReference:newReference callback:^(BOOL success, NSError *error) {
            if(error == nil && success){
                dispatch_async(dispatch_get_main_queue(), ^{
                    [spinner stopAnimating];
                    [self.navigationController popViewControllerAnimated:YES];
                });
            }else{
                dispatch_async(dispatch_get_main_queue(), ^{
                    [spinner stopAnimating];
                    NSString *errorMessage = (error) ? [@"Add Reference failed. Reason: " stringByAppendingString: error.description] : @"Invalid Url";
                    UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:errorMessage delegate:self cancelButtonTitle:@"Ok" otherButtonTitles:nil, nil];
                    [alert show];
                });
            }
        }];
        [task resume];
    }else{
        dispatch_async(dispatch_get_main_queue(), ^{
            UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:@"Complete all fields" delegate:self cancelButtonTitle:@"Ok" otherButtonTitles:nil, nil];
            [alert show];
        });
    }
}
    ```

    And add the import sentence
    ```
    #import "ProjectClient.h"
    ```

02. On **ProjectDetailsViewController.m** uncomment this line in the method **prepareForSegue:sender:**

    ```
    //controller.project = self.project;
    ```

03. Back in **CreateReferenceViewController.h**, add the variable

    ```
    @property ListItem* project;
    ```

    And add the import
    ```
    #import "office365-lists-sdk/ListItem.h"
    ```

04. Build and Run the app, and check everything is ok. Now you can add a reference to a project

    ![](img/fig.21.png)

###Task6 - Wiring up ReferenceDetailsView

01. On **ReferenceDetailsViewController.m** add the initialization method

    ```
    - (void)viewDidLoad
{
    [super viewDidLoad];
    
    [self.navigationController.navigationBar setBackgroundImage:nil
                                                  forBarMetrics:UIBarMetricsDefault];
    self.navigationController.navigationBar.shadowImage = nil;
    self.navigationController.navigationBar.translucent = NO;
    self.navigationController.view.backgroundColor = nil;
    
    NSDictionary *dic =[self.selectedReference getData:@"URL"];
    
    if(![[self.selectedReference getData:@"Comments"] isEqual:[NSNull null]]){
        self.descriptionLbl.text = [self.selectedReference getData:@"Comments"];
    }else{
        self.descriptionLbl.text = @"";
    }
    self.urlTableCell.scrollEnabled = NO;
    self.navigationItem.title = [dic valueForKey:@"Description"];
}
    ```

02. Add the table actions

    ```
    - (NSInteger)tableView:(UITableView *)tableView numberOfRowsInSection:(NSInteger)section{
    return 1;
}
- (UITableViewCell *)tableView:(UITableView *)tableView cellForRowAtIndexPath:(NSIndexPath *)indexPath{
    NSString* identifier = @"referenceDetailsTableCell";
    ReferenceDetailTableCellTableViewCell *cell =[tableView dequeueReusableCellWithIdentifier: identifier ];
    
    NSDictionary *dic =[self.selectedReference getData:@"URL"];
    
    cell.urlContentLBL.text = [dic valueForKey:@"Url"];
    
    return cell;
}
- (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath
{
    NSDictionary *dic =[self.selectedReference getData:@"URL"];
    NSURL *url = [NSURL URLWithString:[dic valueForKey:@"Url"]];
    
    if (![[UIApplication sharedApplication] openURL:url]) {
        NSLog(@"%@%@",@"Failed to open url:",[url description]);
    }
}
    ```

03. Forward navigation

    ```
    - (void)prepareForSegue:(UIStoryboardSegue *)segue sender:(id)sender{
    if ([segue.identifier isEqualToString:@"editReference"]){
        EditReferenceViewController *controller = (EditReferenceViewController *)segue.destinationViewController;
        controller.token = self.token;
        //controller.selectedReference = self.selectedReference;
    }
}
    ```

04. On **ReferenceDetailsViewController.h**, add the variable:

    ```
    @property ListItem* selectedReference;
    ```

    And the import sentence
    ```
    #import "office365-lists-sdk/ListItem.h"
    ```

05. On the **ProjectDetailsViewController.m** uncomment this line on the method **prepareForSegue:sender:**

    ```
    //controller.selectedReference = self.selectedReference;
    ```

06. Build and Run the app, and check everything is ok. Now you can see the Reference details.

    ![](img/fig.22.png)


###Task7 - Wiring up EditReferenceView

01. On **ReferenceDetailsViewController.m** uncomment this line on the method **prepareForSegue:sender**

    ```
    controller.selectedReference = self.selectedReference;
    ```

02. On **EditReferenceViewController.h** add a variable:

    ```
    @property ListItem* selectedReference;
    ```

    And the import sentence
    ```
    #import "office365-lists-sdk/ListItem.h"
    ```

03. On the **EditReferenceViewController.m**, change the **viewDidLoad** method

    ```
    - (void)viewDidLoad
{
    [super viewDidLoad];
    
    [self.navigationController.navigationBar setBackgroundImage:nil
                                                  forBarMetrics:UIBarMetricsDefault];
    self.navigationController.navigationBar.shadowImage = nil;
    self.navigationController.navigationBar.translucent = NO;
    self.navigationController.title = @"Edit Reference";
    
    self.navigationController.view.backgroundColor = nil;
    
    NSDictionary *dic =[self.selectedReference getData:@"URL"];
    
    self.referenceUrlTxt.text = [dic valueForKey:@"Url"];
    
    if(![[self.selectedReference getData:@"Comments"] isEqual:[NSNull null]]){
        self.referenceDescription.text = [self.selectedReference getData:@"Comments"];
    }else{
        self.referenceDescription.text = @"";
    }

    self.referenceTitle.text = [dic valueForKey:@"Description"];
}
    ```

04. Change the **updateReference** method

    ```
    -(void)updateReference{
    if((![self.referenceUrlTxt.text isEqualToString:@""]) && (![self.referenceDescription.text isEqualToString:@""]) && (![self.referenceTitle.text isEqualToString:@""])){
        UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
        spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
        [self.view addSubview:spinner];
        spinner.hidesWhenStopped = YES;
        
        [spinner startAnimating];
        
        
        ListItem* editedReference = [[ListItem alloc] init];
        
        NSDictionary* urlDic = [NSDictionary dictionaryWithObjects:@[self.referenceUrlTxt.text, self.referenceTitle.text] forKeys:@[@"Url",@"Description"]];
        
        NSDictionary* dic = [NSDictionary dictionaryWithObjects:@[urlDic, self.referenceDescription.text, [self.selectedReference getData:@"Project"], self.selectedReference.Id] forKeys:@[@"URL",@"Comments",@"Project",@"Id"]];
        
        [editedReference initWithDictionary:dic];
        

        ProjectClient* client = [ProjectClient getClient:self.token];
        
        NSURLSessionTask* task = [client updateReference:editedReference callback:^(BOOL result, NSError *error) {
            if(error == nil && result){
                dispatch_async(dispatch_get_main_queue(), ^{
                    [spinner stopAnimating];
                    ProjectDetailsViewController *View = [self.navigationController.viewControllers objectAtIndex:self.navigationController.viewControllers.count-3];
                    [self.navigationController popToViewController:View animated:YES];
                });
            }else{
                dispatch_async(dispatch_get_main_queue(), ^{
                    [spinner stopAnimating];
                    NSString *errorMessage = (error) ? [@"Update Reference failed. Reason: " stringByAppendingString: error.description] : @"Invalid Url";
                    UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:errorMessage delegate:self cancelButtonTitle:@"Ok" otherButtonTitles:nil, nil];
                    [alert show];
                });
            }
        }];
        [task resume];
    }else{
        dispatch_async(dispatch_get_main_queue(), ^{
            UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:@"Complete all fields" delegate:self cancelButtonTitle:@"Ok" otherButtonTitles:nil, nil];
            [alert show];
        });
    }
}
    ```

05. Change the **deleteReference**

    ```
    -(void)deleteReference{
    UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    
    [spinner startAnimating];
    
    ProjectClient* client = [ProjectClient getClient:self.token];
    
    NSURLSessionTask* task = [client deleteListItem:@"Research References" itemId:self.selectedReference.Id callback:^(BOOL result, NSError *error) {
        if(error == nil){
            dispatch_async(dispatch_get_main_queue(), ^{
                [spinner stopAnimating];
                ProjectDetailsViewController *View = [self.navigationController.viewControllers objectAtIndex:self.navigationController.viewControllers.count-3];
                [self.navigationController popToViewController:View animated:YES];
            });
        }else{
            NSString *errorMessage = [@"Delete Reference failed. Reason: " stringByAppendingString: error.description];
            UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error" message:errorMessage delegate:self cancelButtonTitle:@"Retry" otherButtonTitles:@"Cancel", nil];
            [alert show];
        }
    }];
    
    [task resume];
}
    ```

06. Add the import sentence to the **ProjectClient** class

    ```
    #import "ProjectClient.h"
    ```

07. Build and Run the app, and check everything is ok. Now you can edit and delete a reference.

    ![](img/fig.23.png)


###Task8 - Wiring up Add Reference Safari Extension

```    
The app provides a Safari action extension, that allows the user to share a url and add it to
a project using a simple screen, without entering the main app.
```

01. Add the **loadData** method body on **ActionViewController.m**

    ```
    -(void)loadData{
    //Create and add a spinner
    UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    [spinner startAnimating];
    
    ProjectClientEx* client = [ProjectClientEx getClient:self.token];
    
    NSURLSessionTask* task = [client getList:@"Research Projects" callback:^(ListEntity *list, NSError *error) {
        
        //If list doesn't exists, create one with name Research Projects
        if(list){
            dispatch_async(dispatch_get_main_queue(), ^{
                [self getProjectsFromList:spinner];
            });
        }else{
            dispatch_async(dispatch_get_main_queue(), ^{
                [self createProjectList:spinner];
            });
        }
        
    }];
    [task resume];
}
    ```

    And add the import sentence
    ```
    #import "ProjectClientEx.h"
    ```


02. Load Projects from the list

    ```
    -(void)getProjectsFromList:(UIActivityIndicatorView *) spinner{
    ProjectClientEx* client = [ProjectClientEx getClient:self.token];
    
    NSURLSessionTask* listProjectsTask = [client getListItems:@"Research Projects" callback:^(NSMutableArray *listItems, NSError *error) {
        if(!error){
            self.projectsList = listItems;
            
            dispatch_async(dispatch_get_main_queue(), ^{
                [self.projectTable reloadData];
                [spinner stopAnimating];
            });
        }
    }];
    [listProjectsTask resume];
}
    ```

03. Create the List if not exists

    ```
    -(void)createProjectList:(UIActivityIndicatorView *) spinner{
    ProjectClientEx* client = [ProjectClientEx getClient:self.token];
    
    ListEntity* newList = [[ListEntity alloc ] init];
    [newList setTitle:@"Research Projects"];
    
    NSURLSessionTask* createProjectListTask = [client createList:newList :^(ListEntity *list, NSError *error) {
        [spinner stopAnimating];
    }];
    [createProjectListTask resume];
}
    ```

04. Finally add the table actions and events, including the selection and the references sharing

    ```
    - (UITableViewCell *)tableView:(UITableView *)tableView cellForRowAtIndexPath:(NSIndexPath *)indexPath
{
    NSString* identifier = @"ProjectListCell";
    ProjectTableExtensionViewCell *cell =[tableView dequeueReusableCellWithIdentifier: identifier ];
    
    ListItem *item = [self.projectsList objectAtIndex:indexPath.row];
    cell.ProjectName.text = [item getTitle];
    
    return cell;
}
- (NSInteger)tableView:(UITableView *)tableView numberOfRowsInSection:(NSInteger)section
{
    return [self.projectsList count];
}
- (CGFloat)tableView:(UITableView *)tableView heightForRowAtIndexPath:(NSIndexPath *)indexPath{
    return 40;
}
- (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath
{
    UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(135,140,50,50)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    
    [spinner startAnimating];
    
    currentEntity= [self.projectsList objectAtIndex:indexPath.row];
    
    NSString* obj = [NSString stringWithFormat:@"{'Url':'%@', 'Description':'%@'}", self.urlTxt.text, @""];
    NSDictionary* dic = [NSDictionary dictionaryWithObjects:@[obj, @"", [NSString stringWithFormat:@"%@", currentEntity.Id]] forKeys:@[@"URL", @"Comments", @"Project"]];
    
    ListItem* newReference = [[ListItem alloc] initWithDictionary:dic];
    
    __weak ActionViewController *sself = self;
    
    NSURLSessionTask* task =[[ProjectClientEx getClient:self.token] addReference:newReference callback:^(BOOL success, NSError *error) {
        if(error == nil){
            dispatch_async(dispatch_get_main_queue(), ^{
                sself.projectTable.hidden = true;
                sself.selectProjectLbl.hidden = true;
                sself.successMsg.hidden = false;
                sself.successMsg.text = [NSString stringWithFormat:@"Reference added successfully to the %@ Project.", [currentEntity getTitle]];
                [spinner stopAnimating];
            });
        }
    }];
    
    [task resume];
}
    ```

05. To Run the app, you should select the correct target. To do so, follow the steps:

    On the **Run/Debug panel control**, you will see the target selected
    ![](img/fig.26.png)

    Click on the target name and select the **Extension Target** and an iOS simulator
    ![](img/fig.27.png)

    Now you can Build and Run the application, but first we have to select what native application
    will open in order to access the extension. In this case, we select **Safari**                                    
    ![](img/fig.28.png)


06. Build and Run the application, check everything is ok. Now you can share a reference url from safari and attach it to a Project with our application.

    Custom Action Extension                                                                                       
    ![](img/fig.24.png)

    Simple view to add a Reference to a Project                                                                  
    ![](img/fig.25.png)

##Summary

By completing this hands-on lab you have learnt:

01. The way to connect an iOS application with an Office365 tenant.

02. How to retrieve information from Sharepoint lists.

03. How to handle the responses in JSON format. And communicate with the infrastructure.

