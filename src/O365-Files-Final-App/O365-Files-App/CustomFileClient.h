#import <office365_drive_sdk/office365_drive_sdk.h>

@interface CustomFileClient : NSObject
/*- (NSURLSessionDataTask *)getFiles:(NSString *)folder callback :(void (^)(NSMutableArray *files, NSError *))callback;
- (NSURLSessionDataTask *)download:(NSString *)fileName callback :(void (^)(NSData *data, NSError *error))callback;*/
+(MSSharePointClient*)getClient:(NSString *) token;
@end
