#import <office365_drive_sdk/office365_drive_sdk.h>
#import "ViewController.h"

@interface FileDetailsViewController : ViewController<NSURLConnectionDelegate>

{
    NSMutableData *_responseData;
}

@property NSString *token;
@property MSSharePointItem *file;
@property (weak, nonatomic) IBOutlet UIBarButtonItem *downloadButton;
@property (nonatomic, strong) UIDocumentInteractionController *docInteractionController;
@property (weak, nonatomic) IBOutlet UILabel *fileName;
@property (weak, nonatomic) IBOutlet UILabel *lastModified;
@property (weak, nonatomic) IBOutlet UILabel *created;



@end
