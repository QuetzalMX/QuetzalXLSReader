//
//  QZCell.h
//  QZXLSReader
//
//  Created by Fernando Olivares on 10/27/13.
//  Copyright 2013 Fernando Olivares.
//

#import "QZWorkbook.h"

typedef struct xlsCell cell;

typedef enum {
    Blank = 0,
    String,
    Integer,
    Float,
    Bool,
    Error,
    Unknown
} ContentType;

@interface QZCell : NSObject

@property (nonatomic, assign, readonly) ContentType type;
@property (nonatomic, assign, readonly) QZLocation location;
@property (nonatomic, strong, readonly) id content;

- (id)initWithContent:(struct xlsCell *)cell;

@end