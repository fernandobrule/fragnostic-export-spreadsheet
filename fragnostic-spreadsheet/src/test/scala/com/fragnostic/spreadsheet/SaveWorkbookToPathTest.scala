package com.fragnostic.spreadsheet

import java.io.File

import com.fragnostic.spreadsheet.CakeServiceSpreadsheet.spreadsheet
import com.fragnostic.spreadsheet.support.AbstractTest
import org.apache.poi.ss.usermodel.Workbook

class SaveWorkbookToPathTest extends AbstractTest {

  describe("Save Workbook To Path Test") {

    it("Can Save Workbook To Path") {

      val wb: Workbook = spreadsheet.getWorkbook(pathXls) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      val newPath: String = s"$pathXls---copy.xls"

      val success: String = spreadsheet.save(wb, newPath) fold (
        error => throw new IllegalStateException(error),
        success => success)

      success should be("spreadsheet.service.save.workbook.success")

      new File(newPath).exists should be(true)

    }

    it("Can Not Save Workbook To Path That Does Not Exists") {

      val wb: Workbook = spreadsheet.getWorkbook(pathXls) fold (
        error => throw new IllegalStateException(error),
        wb => wb)

      val newPath: String = s"dsfsdfsdfdsfsdfsdf/asdfsf/sdfs"

      val error: String = spreadsheet.save(wb, newPath) fold (
        error => error,
        success => success)

      error should be("spreadsheet.service.save.workbook.fail")

    }

  }

}
