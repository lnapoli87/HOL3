#import "ViewController.h"

@interface FileListViewController : ViewController
@property (weak, nonatomic) IBOutlet UITableView *tableView;
@property NSString *token;
@property NSMutableArray *files;
@end
