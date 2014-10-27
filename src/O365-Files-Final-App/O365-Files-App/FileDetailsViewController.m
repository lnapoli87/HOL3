//
//  FileDetailsViewController.m
//  O365-Files-App
//
//  Created by Lucas Damian Napoli on 24/10/14.
//  Copyright (c) 2014 MS Open Tech. All rights reserved.
//

#import "FileDetailsViewController.h"
#import "CustomFileClient.h"
#import "office365-base-sdk/HttpConnection.h"
#import "office365-base-sdk/Credentials.h"

@interface FileDetailsViewController ()

@end

@implementation FileDetailsViewController

UIActivityIndicatorView* spinner;

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
    
    self.fileName.text = self.file.Name;
    self.lastModified.text = [self.file.TimeLastModified substringToIndex:10];
    self.created.text = [self.file.TimeCreated substringToIndex:10];

    
    [self loadFile];
}

- (void) loadFile{
    double x = ((self.navigationController.view.frame.size.width) - 20)/ 2;
    double y = ((self.navigationController.view.frame.size.height) - 150)/ 2;
    spinner = [[UIActivityIndicatorView alloc]initWithFrame:CGRectMake(x, y, 20, 20)];
    spinner.activityIndicatorViewStyle = UIActivityIndicatorViewStyleGray;
    [self.view addSubview:spinner];
    spinner.hidesWhenStopped = YES;
    [spinner startAnimating];
    
    NSString *fileUrlString = self.file.Url;
    
    CustomFileClient *client = [CustomFileClient getClient:self.token];
    
    NSURLSessionDataTask *task = [client download:self.file.Name callback:^(NSData *data, NSError *error) {
        if ( data )
        {
            NSArray       *paths = NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES);
            NSString  *documentsDirectory = [paths objectAtIndex:0];
            
            NSString  *filePath = [NSString stringWithFormat:@"%@/%@", documentsDirectory,self.file.Name];
            [data writeToFile:filePath atomically:YES];
            
            NSURL *fileUrl = [NSURL fileURLWithPath:filePath];
            
            self.docInteractionController = [UIDocumentInteractionController interactionControllerWithURL:fileUrl];
            self.docInteractionController.delegate = self;
        }
        dispatch_async(dispatch_get_main_queue(), ^{
            [spinner stopAnimating];
        });
    }];
    
    [task resume];
}

- (UIViewController *) documentInteractionControllerViewControllerForPreview: (UIDocumentInteractionController *) controller
{
    return [self navigationController];
}

- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
}


- (IBAction)downloadAction:(id)sender {
    [self.docInteractionController presentPreviewAnimated:YES];
}



@end
