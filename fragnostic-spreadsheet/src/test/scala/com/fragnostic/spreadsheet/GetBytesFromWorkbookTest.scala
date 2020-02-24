package com.fragnostic.spreadsheet

import better.files.{ File => BFFile }
import com.fragnostic.spreadsheet.CakeServiceSpreadsheet.spreadsheet
import com.fragnostic.spreadsheet.support.AbstractTest
import org.apache.poi.ss.usermodel.Workbook

class GetBytesFromWorkbookTest extends AbstractTest {

  describe("Get Bytes From Workbook Test") {

    it("Can Get Bytes From Workbook") {

      val wb: Workbook = spreadsheet.getWorkbook(pathXls) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      val bytes: Array[Byte] = spreadsheet.getBytes(wb) fold (
        error => throw new IllegalStateException(error),
        bytes => bytes)

      bytes.length should be(BFFile(pathXls).loadBytes.length)

    }

  }

}
