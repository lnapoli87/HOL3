//
//  CustomFileClient.h
//  O365-Files-App
//
//  Created by Lucas Damian Napoli on 24/10/14.
//  Copyright (c) 2014 MS Open Tech. All rights reserved.
//

#import "office365-files-sdk/FileClient.h"

@interface CustomFileClient : FileClient
- (NSURLSessionDataTask *)getFiles:(NSString *)folder callback :(void (^)(NSMutableArray *files, NSError *))callback;
- (NSURLSessionDataTask *)download:(NSString *)fileName callback :(void (^)(NSData *data, NSError *error))callback;
+(FileClient*)getClient:(NSString *) token;
@end
