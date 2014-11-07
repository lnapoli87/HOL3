#import "FileListViewController.h"
#import "FileListCellTableViewCell.h"
#import "CustomFileClient.h"
#import "FileDetailsViewController.h"

@implementation FileListViewController

NSDateFormatter *formatter;
MSSharePointItem* currentEntity;

- (void)viewDidLoad {
    [super viewDidLoad];
    
    [self.navigationController.navigationBar setBackgroundImage:nil
                                                  forBarMetrics:UIBarMetricsDefault];
    self.navigationController.navigationBar.shadowImage = nil;
    self.navigationController.navigationBar.translucent = NO;
    self.navigationController.view.tintColor = [UIColor colorWithRed:226.0/255.0 green:37.0/255.0 blue:7.0/255.0 alpha:1];
    self.navigationController.navigationBar.tintColor = [UIColor whiteColor];
    self.navigationController.navigationBar.barTintColor = [UIColor colorWithRed:226.0/255.0 green:37.0/255.0 blue:7.0/255.0 alpha:1];
    self.navigationController.navigationBar.titleTextAttributes = [NSDictionary dictionaryWithObjectsAndKeys:
                                                                   [UIColor whiteColor], NSForegroundColorAttributeName, nil];
    

    [[UIApplication sharedApplication] setStatusBarStyle:UIStatusBarStyleLightContent];
    
    formatter = [[NSDateFormatter alloc] init];
    [formatter setDateFormat:@"MM-dd-yyyy"];
}

- (void)viewWillAppear:(BOOL)animated{
    if (!self.currentFolder){
        self.navigationController.title = @"File List";
        [self loadData];
    }else{
        self.navigationController.title = self.currentFolder.name;
        [self loadCurrentFolder];
    }
    currentEntity = nil;
}

-(void) loadData{
    //Create and add a spinner
    double x = ((self.navigationController.view.frame.size.width) - 20)/ 2;
    double y = ((self.navigationController.view.frame.size.height) - 150)/ 2;
    UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(x, y, 20, 20)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    [spinner startAnimating];
    
    MSSharePointClient *client = [CustomFileClient getClient:self.token];
    NSURLSessionDataTask *task = [[client getfiles]read:^(NSArray<MSSharePointItem> *files, NSError *error) {
        self.files = files;
        dispatch_async(dispatch_get_main_queue(), ^{
            [self.tableView reloadData];
            [spinner stopAnimating];
        });
    }];
    
    
    [task resume];
}


-(void) loadCurrentFolder{
    //Create and add a spinner
    double x = ((self.navigationController.view.frame.size.width) - 20)/ 2;
    double y = ((self.navigationController.view.frame.size.height) - 150)/ 2;
    UIActivityIndicatorView* spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(x, y, 20, 20)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    [spinner startAnimating];
    
    MSSharePointClient *client = [CustomFileClient getClient:self.token];
    
    [[[[[[client getfiles] getById:self.currentFolder.id] asFolder] getchildren] read:^(NSArray<MSSharePointItem> *files, NSError *error) {
        self.files = files;
        dispatch_async(dispatch_get_main_queue(), ^{
            [self.tableView reloadData];
            [spinner stopAnimating];
        });
    }] resume];
}



-(void) viewWillDisappear:(BOOL)animated {
    if ([self.navigationController.viewControllers indexOfObject:self]==NSNotFound && !self.currentFolder) {
        [self.navigationController.navigationBar setBackgroundImage:[UIImage new]
                                                      forBarMetrics:UIBarMetricsDefault];
        self.navigationController.navigationBar.shadowImage = [UIImage new];
        self.navigationController.navigationBar.translucent = YES;
        self.navigationController.view.backgroundColor = [UIColor clearColor];
        [[UIApplication sharedApplication] setStatusBarStyle:UIStatusBarStyleDefault];
    }
    [super viewWillDisappear:animated];
}

- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
}

- (NSInteger)tableView:(UITableView *)tableView numberOfRowsInSection:(NSInteger)section{
    return self.files.count;
}

- (UITableViewCell *)tableView:(UITableView *)tableView cellForRowAtIndexPath:(NSIndexPath *)indexPath{
    NSString* identifier = @"fileListCell";
    FileListCellTableViewCell *cell =[tableView dequeueReusableCellWithIdentifier: identifier ];
    
    MSSharePointItem *file = [self.files objectAtIndex:indexPath.row];
    
    NSString *lastModifiedString = [formatter stringFromDate:file.dateTimeLastModified];
    
    cell.fileName.text = file.name;
    cell.lastModified.text = [NSString stringWithFormat:@"Last modified on %@", lastModifiedString];
    
    return cell;
}

- (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath{
    currentEntity= [self.files objectAtIndex:indexPath.row];
    
    if ([currentEntity.type isEqualToString:@"Folder"]){
        FileListViewController *controller = [[UIStoryboard storyboardWithName:@"Main" bundle:nil] instantiateViewControllerWithIdentifier:@"fileList"];
        controller.token = self.token;
        controller.currentFolder = currentEntity;
        
        [self.navigationController pushViewController:controller animated:YES];
    }else{
        [self performSegueWithIdentifier:@"detail" sender:self];
    }
}
- (BOOL)shouldPerformSegueWithIdentifier:(NSString *)identifier sender:(id)sender{
    return ([identifier isEqualToString:@"detail"] && currentEntity);
}

-(void) prepareForSegue:(UIStoryboardSegue *)segue sender:(id)sender{
    if([segue.identifier isEqualToString:@"detail"]){
        FileDetailsViewController *ctrl = (FileDetailsViewController *)segue.destinationViewController;
        ctrl.token = self.token;
        ctrl.file = currentEntity;
    }
}

@end
