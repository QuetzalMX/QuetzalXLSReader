//
//  QZWorkSheet.h
//  QZXLSReader
//
//  Created by Fernando Olivares on 10/27/13.
//  Copyright 2013 Fernando Olivares.
//

#import <Foundation/Foundation.h>

@class QZCell;
typedef struct xlsWorkSheet worksheet;
typedef struct QZLocation QZLocation;

@interface QZWorkSheet : NSObject

@property (readonly) BOOL isOpen;
@property (nonatomic, strong) NSString *name;
@property (nonatomic, strong) NSArray *rows;
@property (nonatomic, strong) NSArray *columns;

- (id)initFromXLSWorkSheet:(struct xlsWorkSheet *)workSheet
                 withIndex:(NSUInteger)sheetIndex;

- (void)open;
- (void)close;
- (QZCell *)cellAtPoint:(QZLocation)location;

- (NSArray *)arrayRepresentationOfSimpleWorkSheet;

@end