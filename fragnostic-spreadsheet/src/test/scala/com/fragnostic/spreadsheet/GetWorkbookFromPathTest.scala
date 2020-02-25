package com.fragnostic.spreadsheet

import com.fragnostic.spreadsheet.CakeServiceSpreadsheet.spreadsheet

import com.fragnostic.spreadsheet.support.AbstractTest
import org.apache.poi.ss.usermodel.Workbook

class GetWorkbookFromPathTest extends AbstractTest {

  describe("Get Workbook From Path Test") {

    it("Can Get Wokbook from Path") {

      val wb: Workbook = spreadsheet.getWorkbook(pathXls) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue should be("Air intake manifold")

    }
  }

}