//  QZWorkbook.m
//  QZXLSReader
//
//  Created by Fernando Olivares on 10/27/13.
//  Copyright 2013 Fernando Olivares.
//

#import "QZWorkbook.h"
#import "QZWorkSheet.h"
#import "xls.h"

@interface QZWorkbook (){
	xlsWorkBook *_workBook;
}

@end

@implementation QZWorkbook

@synthesize createdBy = _createdBy;
@synthesize excelVersion = _excelVersion;
@synthesize fileName = _fileName;
@synthesize modifiedBy = _modifiedBy;
@synthesize workSheets = _workSheets;

- (id)initWithContentsOfXLS:(NSURL *)filePathURL;
{
    //Validate your inputs and self.
    if(self != [super init] || ![filePathURL isFileURL])
        return nil;

    //Get the filePath and attempt to open it.
	if(![self openXLSFile:filePathURL])
        return nil;
    
	return self;
}

- (BOOL)openXLSFile:(NSURL *)filePathURL;
{
    //Try to open the raw XLS file.
    const char *filePath = [[filePathURL path] cStringUsingEncoding:NSUTF8StringEncoding];
    _workBook = xls_open(filePath, "UTF-8");
    
    //Did it work?
    if(!_workBook)
        return NO;
    
    //It worked.
    _fileName = [filePathURL lastPathComponent];
    
    //Parse it.
	xls_parseWorkBook(_workBook);
	xlsSummaryInfo *summary = xls_summaryInfo(_workBook);
    
    //Get the info out of it.
    _excelVersion = [NSString stringWithCString:(char *)summary->appName encoding:NSUTF8StringEncoding];
    _createdBy = [NSString stringWithCString:(char *)summary->author encoding:NSUTF8StringEncoding];
    _modifiedBy = [NSString stringWithCString:(char *)summary->lastAuthor encoding:NSUTF8StringEncoding];
    xls_close_summaryInfo(summary);
    
    //Parse its worksheets.
    NSMutableArray *workSheets = [NSMutableArray new];
    for(NSInteger sheetCounter = 0; sheetCounter < _workBook->sheets.count; sheetCounter++){
        xlsWorkSheet *workSheet = xls_getWorkSheet(_workBook, sheetCounter);
        QZWorkSheet *newWorkSheet = [[QZWorkSheet alloc] initFromXLSWorkSheet:workSheet
                                                            withIndex:sheetCounter];
        
        if(newWorkSheet)
            [workSheets addObject:newWorkSheet];
    }
    
    _workSheets = [NSArray arrayWithArray:workSheets];
    
    return YES;
}

- (void)close;
{
    //Frees the memory.
	xls_close_WB(_workBook);
}

- (NSString *)description;
{
    return [NSString stringWithFormat:@"%@ DHXLSWorkbook: %@ by %@. Last modified by: %@ in %@.", [super description], self.fileName, self.createdBy, self.modifiedBy, self.excelVersion];
}

@end