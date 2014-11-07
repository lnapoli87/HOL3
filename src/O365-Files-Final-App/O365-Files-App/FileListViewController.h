#import "ViewController.h"
#import <office365_drive_sdk/office365_drive_sdk.h>

@interface FileListViewController : ViewController
@property (weak, nonatomic) IBOutlet UITableView *tableView;
@property NSString *token;
@property NSArray<MSSharePointItem> *files;
@property MSSharePointItem* currentFolder;
@end
