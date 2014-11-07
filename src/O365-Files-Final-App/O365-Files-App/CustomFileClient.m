#import "CustomFileClient.h"

@implementation CustomFileClient

const NSString *apiUrl = @"/_api/files";

/*- (NSURLSessionDataTask *)download:(NSString *)fileName callback :(void (^)(NSData *data, NSError *error))callback{
    
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
}*/
    
+(MSSharePointClient*)getClient:(NSString *) token{
    NSString *url = [NSString alloc];
    NSString* plistPath = [[NSBundle mainBundle] pathForResource:@"Auth" ofType:@"plist"];
    url = [[NSDictionary dictionaryWithContentsOfFile:plistPath] objectForKey:@"o365SharepointTenantUrl"];
    
    MSDefaultDependencyResolver* resolver = [MSDefaultDependencyResolver alloc];
    MSOAuthCredentials* credentials = [MSOAuthCredentials alloc];
    [credentials addToken: token];
    
    MSCredentialsImpl* credentialsImpl = [MSCredentialsImpl alloc];
    
    [credentialsImpl setCredentials:credentials];
    [resolver setCredentialsFactory:credentialsImpl];
    
    return [[MSSharePointClient alloc] initWitUrl:[url stringByAppendingString:@"/_api/v1.0/me"] dependencyResolver:resolver];
}

@end
