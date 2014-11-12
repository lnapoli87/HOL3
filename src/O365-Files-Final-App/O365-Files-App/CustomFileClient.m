#import "CustomFileClient.h"

@implementation CustomFileClient

const NSString *apiUrl = @"/_api/files";

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
