//
//  FileListViewController.m
//  O365-Files-App
//
//  Created by Lucas Damian Napoli on 24/10/14.
//  Copyright (c) 2014 MS Open Tech. All rights reserved.
//

#import "FileListViewController.h"
#import "FileListCellTableViewCell.h"
#import "office365-files-sdk/FileClient.h"
#import "office365-base-sdk/OAuthentication.h"
#import "office365-files-sdk/FileEntity.h"
#import "CustomFileClient.h"
#import "FileDetailsViewController.h"

@implementation FileListViewController

FileEntity* currentEntity;

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
     self.navigationController.title = @"File List";
    
    self.files = [[NSMutableArray alloc] init];
    currentEntity = nil;
    
}

- (void)viewWillAppear:(BOOL)animated{
    [self loadData];
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



-(void) viewWillDisappear:(BOOL)animated {
    if ([self.navigationController.viewControllers indexOfObject:self]==NSNotFound) {
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
    
    FileEntity *file = [self.files objectAtIndex:indexPath.row];
    
    cell.fileName.text = file.Name;
    cell.lastModified.text = [NSString stringWithFormat:@"Last modified on %@", [file.TimeLastModified substringToIndex:10]];
    
    return cell;
}

- (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath{
    currentEntity= [self.files objectAtIndex:indexPath.row];
    
    [self performSegueWithIdentifier:@"detail" sender:self];
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
