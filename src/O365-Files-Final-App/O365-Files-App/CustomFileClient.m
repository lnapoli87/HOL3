#import "CustomFileClient.h"
#import "office365-base-sdk/NSString+NSStringExtensions.h"
#import "office365-base-sdk/HttpConnection.h"
#import "office365-base-sdk/Constants.h"
#import "office365-base-sdk/OAuthentication.h"

@implementation CustomFileClient

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

- (NSURLSessionDataTask *)download:(NSString *)fileName callback :(void (^)(NSData *data, NSError *error))callback{
    
    NSString *url = [NSString stringWithFormat:@"%@%@('%@')/download", self.Url , apiUrl, [fileName stringByAddingPercentEscapesUsingEncoding:NSUTF8StringEncoding]];

    HttpConnection *connection = [[HttpConnection alloc] initWithCredentials:self.Credential url:url ];

    NSString *method = (NSString*)[[Constants alloc] init].Method_Get;

    return [connection execute:method callback:^(NSData  *data, NSURLResponse *reponse, NSError *error) {
        FileEntity *file = [[FileEntity alloc] init];
        
        NSDictionary *jsonResult = [NSJSONSerialization JSONObjectWithData:data
                                                                   options: NSJSONReadingMutableContainers
                                                                     error:nil];
        
        NSDictionary *jsonArray = [jsonResult valueForKey : @"d"];
        
        if(error == nil){
            [file createFromJson: jsonArray];
        }
        
        callback(data, error);
    }];
}
    
+(CustomFileClient*)getClient:(NSString *) token{
    OAuthentication* authentication = [OAuthentication alloc];
    [authentication setToken:token];
    
    NSString *url = [NSString alloc];
    NSString* plistPath = [[NSBundle mainBundle] pathForResource:@"Auth" ofType:@"plist"];
    url = [[NSDictionary dictionaryWithContentsOfFile:plistPath] objectForKey:@"o365SharepointTenantUrl"];
    
    
    return [[CustomFileClient alloc] initWithUrl: url credentials: authentication];
}

@end
