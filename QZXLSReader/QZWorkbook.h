//
//  QZWorkbook.h
//  QZXLSReader
//
//  Created by Fernando Olivares on 10/27/13.
//  Copyright 2013 Fernando Olivares.
//

#import <Foundation/Foundation.h>

struct QZLocation {
    NSUInteger row;
    NSUInteger column;
};
typedef struct QZLocation QZLocation;

@interface QZWorkbook : NSObject

- (id)initWithContentsOfXLS:(NSURL *)filePathURL;

- (void)close;

@property (nonatomic, strong) NSArray *workSheets;
@property (nonatomic, strong, readonly) NSString *fileName;
@property (nonatomic, strong, readonly) NSString *excelVersion;
@property (nonatomic, strong, readonly) NSString *createdBy;
@property (nonatomic, strong, readonly) NSString *modifiedBy;

@end