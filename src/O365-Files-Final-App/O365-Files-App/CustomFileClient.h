#import <office365_drive_sdk/office365_drive_sdk.h>

@interface CustomFileClient : NSObject
+(MSSharePointClient*)getClient:(NSString *) token;
@end
