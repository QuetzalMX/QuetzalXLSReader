//
//  QZCell.m
//  QZXLSReader
//
//  Created by Fernando Olivares on 10/27/13.
//  Copyright 2013 Fernando Olivares.
//

#import "xls.h"
#import "QZCell.h"

@implementation QZCell

@synthesize location = _location;
@synthesize type = _type;
@synthesize content = _content;

- (id)initWithContent:(xlsCell *)cell;
{
    if(self != [super init])
        return nil;
    
	_location.row = cell->row;
    _location.column = cell->col;
    
	switch(cell->id) {
        case 0x0006:	//FORMULA
            // test for formula, if
            if(cell->l == 0) {
                _type = Float;
                _content = [NSNumber numberWithDouble:cell->d];
            } else {
                if(!strcmp((char *)cell->str, "bool")) {
                    BOOL b = (BOOL)cell->d;
                    _type = Bool;
                    _content = [NSNumber numberWithBool:b];
                } else
                    if(!strcmp((char *)cell->str, "error")) {
                        // FIXME: Why do we convert the double cell->d to NSInteger?
                        NSInteger err = (NSInteger)cell->d;
                        _type = Error;
                        _content = [NSNumber numberWithInteger:err];
                    } else {
                        _type = String;
                    }
            }
            break;
        case 0x00FD:	//LABELSST
        case 0x0204:	//LABEL
            _type = String;
            _content = [NSString stringWithCString:(char *)cell->str encoding:NSUTF8StringEncoding];
            break;
        case 0x0203:	//NUMBER
        case 0x027E:	//RK
            _type = Float;
            _content = [NSNumber numberWithDouble:cell->d];
            break;
        default:
            _type = Unknown;
            break;
    }
    
    return self;
}

- (NSString *)description;
{
    return [NSString stringWithFormat:@"%@ (%d, %d): %@", super.description, self.location.row, self.location.column, self.content];
}

@end