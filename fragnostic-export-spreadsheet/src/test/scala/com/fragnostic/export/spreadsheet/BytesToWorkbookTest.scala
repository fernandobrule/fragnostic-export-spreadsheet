package com.fragnostic.`export`.spreadsheet

import com.fragnostic.export.spreadsheet.support.BaseTest
import org.apache.poi.ss.usermodel.{ Cell, Row, Sheet, Workbook }

class BytesToWorkbookTest extends BaseTest {

  describe("Bytes To Workbook Test") {
    // https://mkyong.com/java/apache-poi-reading-and-writing-excel-file-in-java/
    it("Can Get Wokbook from Bytes") {

      val path: String = "/Users/fernandobrule/Documents/articulos-58ad5300-4780-4a8a-9477-550abb3a6df9.xls"

      val wb: Workbook = CakeServiceExportSpreadsheet.export.getWorkbook(path) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      val index: Int = 0 // 0-based
      val sheet: Sheet = wb.getSheetAt(index)

      val rowNum: Int = 0 // 0-based
      val row: Row = sheet.getRow(rowNum)

      val cellnum: Int = 0 // 0-based
      val cell: Cell = row.getCell(cellnum)

      println(cell.getStringCellValue)

    }
  }

}