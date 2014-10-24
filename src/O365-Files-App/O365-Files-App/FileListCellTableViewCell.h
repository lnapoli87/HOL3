//
//  FileListCellTableViewCell.h
//  O365-Files-App
//
//  Created by Lucas Damian Napoli on 24/10/14.
//  Copyright (c) 2014 MS Open Tech. All rights reserved.
//

#import <UIKit/UIKit.h>

@interface FileListCellTableViewCell : UITableViewCell
@property (weak, nonatomic) IBOutlet UILabel *fileName;
@property (weak, nonatomic) IBOutlet UILabel *lastModified;

@end
