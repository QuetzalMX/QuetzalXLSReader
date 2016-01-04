//
//  QZXLSWorkbook.swift
//  QZXLSReader
//
//  Created by Fernando Olivares on 1/3/16.
//  Copyright © 2016 Fernando Olivares. All rights reserved.

import Foundation
import CoreGraphics

extension String : ErrorType { }

public class Workbook {
    
    let workBook: UnsafeMutablePointer<xlsWorkBook>
    
    let fileName: String
    let excelVersion: String
    let createdBy: String
    let modifiedBy: String
    
    public init(withLocalXLS filePath: NSURL) throws {
        
        guard let cFilePath = filePath.path?.cStringUsingEncoding(NSUTF8StringEncoding) else {
            workBook = UnsafeMutablePointer()
            fileName = "Workbook1"
            excelVersion = "Unknown"
            createdBy = "Unknown"
            modifiedBy = "Unknown"
            throw "The XLS file path could not be converted to a C string"
        }
        
        workBook = xls_open(cFilePath, "UTF-8".cStringUsingEncoding(NSUTF8StringEncoding))
        
        xls_parseWorkBook(workBook)
        
        let summary = xls_summaryInfo(workBook)
        fileName = filePath.lastPathComponent ?? "Workbook1"
        
        excelVersion = String.fromCString(unsafeBitCast(summary.memory.appName, UnsafePointer<CChar>.self))!
        createdBy = String.fromCString(unsafeBitCast(summary.memory.author, UnsafePointer<CChar>.self))!
        modifiedBy = String.fromCString(unsafeBitCast(summary.memory.lastAuthor, UnsafePointer<CChar>.self))!
        xls_close_summaryInfo(summary)
    }
    
    public func openWorkSheet(atIndex index: UInt) throws -> WorkSheet {
        
        if index > UInt(workBook.memory.sheets.count) {
            throw "Worksheet index out of bounds"
        }
        
        let rawWorkSheet = xls_getWorkSheet(workBook, Int32(index))
        return WorkSheet(withXLSWorkSheet: rawWorkSheet, andIndex: index)
    }
    
    deinit {
        xls_close_WB(workBook);
    }
}

public struct WorkSheet {
    
    let index: UInt
    let name: String
    let rows: [[Cell]]
    
    public init(withXLSWorkSheet workSheet: UnsafeMutablePointer<xlsWorkSheet>, andIndex index: UInt) {
        
        self.index = index
        name = String.fromCString(unsafeBitCast(workSheet.memory.workbook.memory.sheets.sheet.memory.name, UnsafePointer<CChar>.self))!
        
        xls_parseWorkSheet(workSheet)
        
        var allRows: [[Cell]] = []
        for rowIndex in 0...workSheet.memory.rows.lastrow {
            let rawRow = workSheet.memory.rows.row[Int(rowIndex)]
            var rowCells: [Cell] = []
            for cellIndex in 0..<rawRow.cells.count {
                let rawCell = rawRow.cells.cell[Int(cellIndex)]
                let parsedCell = Cell(withRawCell: rawCell)
                
                if parsedCell.position.row == UInt(rowIndex) {
                    rowCells.append(parsedCell)
                }
            }
            
            allRows.append(rowCells)
        }
        
        rows = allRows
        
        xls_close_WS(workSheet);
    }
    
    public subscript(row: Int, column: Int) -> Cell {
        get {
            return rows[row][column]
        }
    }
}

public struct Cell {
    
    let position: CellPosition
    let value: AnyObject
    
    enum CellType : Int {
        case BooleanType, Empty, ErrorType, FloatType, StringType, UnknownType
        
        static func typeFrom(rawCell: st_cell_data) -> CellType {
            let type: CellType
            switch rawCell.id {
            case 513:
                type = Empty
                break
                
            case 0x006:
                if (rawCell.l == 0) {
                    type = FloatType
                } else {
                    guard let stringValue = String.fromCString(unsafeBitCast(rawCell.str, UnsafePointer<CChar>.self)) else { return UnknownType}
                    
                    if stringValue == "bool" {
                        type = BooleanType
                    } else if stringValue == "error" {
                        type = ErrorType
                    } else {
                        type = StringType
                    }

                }
                break
                
            case 253, 516:
                type = StringType
                break
                
            case 515, 638:
                type = FloatType
                break
                
            default:
                type = UnknownType
                break
            }
            
            return type
        }
    }
    
    struct CellPosition {
        let row: UInt
        let column: UInt
    }
    
    init(withRawCell rawCell: st_cell_data) {
        
        position = CellPosition(row: UInt(rawCell.row), column: UInt(rawCell.col))
        switch CellType.typeFrom(rawCell) {
            
        case .FloatType:
            value = Float(rawCell.d)
            break
            
        case .BooleanType:
            value = Bool(rawCell.d)
            break
            
        case .Empty:
            value = ""
            break;
            
        case .ErrorType:
            value = Int(rawCell.d)
            break
            
        case .StringType:
            value = String.fromCString(unsafeBitCast(rawCell.str, UnsafePointer<CChar>.self))!
            break
            
        case .UnknownType:
            value = "(Uknown Value)"
            break
        }
    }
    
}