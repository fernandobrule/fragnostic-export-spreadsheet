package com.fragnostic.spreadsheet

import better.files.{ File => BFFile }
import com.fragnostic.spreadsheet.support.AbstractTest
import org.apache.poi.ss.usermodel.Workbook

class GetWorkbookFromBytesTest extends AbstractTest {

  describe("Get Workbook From Bytes Test") {

    //
    // https://mkyong.com/java/apache-poi-reading-and-writing-excel-file-in-java/
    //
    it("Can Get Wokbook from Bytes") {

      val bytes: Array[Byte] = BFFile(pathXls).byteArray
      println(s"Can Get Wokbook from Bytes - bytes.length:${bytes.length}")

      val wb: Workbook = CakeServiceSpreadsheet.spreadsheet.getWorkbook(bytes) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      wb.getSheetAt(0)
        .getRow(1)
        .getCell(0)
        .getStringCellValue should be("Air intake manifold")

    }
  }

}