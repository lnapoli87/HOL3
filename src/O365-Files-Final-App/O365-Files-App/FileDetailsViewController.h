#import "ViewController.h"
#import "office365-files-sdk/FileEntity.h"

@interface FileDetailsViewController : ViewController<NSURLConnectionDelegate>

{
    NSMutableData *_responseData;
}

@property NSString *token;
@property FileEntity *file;
@property (weak, nonatomic) IBOutlet UIBarButtonItem *downloadButton;
@property (nonatomic, strong) UIDocumentInteractionController *docInteractionController;
@property (weak, nonatomic) IBOutlet UILabel *fileName;
@property (weak, nonatomic) IBOutlet UILabel *lastModified;
@property (weak, nonatomic) IBOutlet UILabel *created;



@end
