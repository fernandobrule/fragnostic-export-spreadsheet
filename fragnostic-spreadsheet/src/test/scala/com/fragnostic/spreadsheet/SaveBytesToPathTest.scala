package com.fragnostic.spreadsheet

import java.io.File

import com.fragnostic.spreadsheet.CakeServiceSpreadsheet.spreadsheet
import com.fragnostic.spreadsheet.support.AbstractTest
import org.apache.poi.ss.usermodel.Workbook

class SaveBytesToPathTest extends AbstractTest {

  describe("Save Bytes To Path Test") {

    it("Can Save Bytes To Path") {

      val wb: Workbook = spreadsheet.getWorkbook(pathXls) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      val bytes: Array[Byte] = spreadsheet.getBytes(wb) fold (
        error => throw new IllegalStateException(error),
        bytes => bytes)

      val newPath: String = s"$pathXls---copy.xls"
      val success: String = spreadsheet.save(bytes, newPath) fold (
        error => throw new IllegalStateException(error),
        success => success)

      success should be("spreadsheet.service.save.bytes.success")

      new File(newPath).exists should be(true)

    }

  }

}
