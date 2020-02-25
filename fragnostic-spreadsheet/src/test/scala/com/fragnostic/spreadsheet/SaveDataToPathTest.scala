package com.fragnostic.spreadsheet

import java.util.Locale

import com.fragnostic.spreadsheet.CakeServiceSpreadsheet.spreadsheet
import com.fragnostic.spreadsheet.support.{ AbstractTest, SomeRow }
import org.apache.poi.ss.usermodel.Row

class SaveDataToPathTest extends AbstractTest {

  describe("Save Data To Path Test") {

    it("Can Save Data To Path") {

      val list = List(
        SomeRow("Pepe", "+56 9 7979 7865"),
        SomeRow("John", "+56 9 3434 5678"))

      val sheetName: String = "pepe"
      val headers: Array[String] = Array("Name", "Telefono")

      def newRow(locale: Locale, someRow: SomeRow, row: Row): Row = {
        row.createCell(0).setCellValue(someRow.name)
        row.createCell(1).setCellValue(someRow.tel)
        row
      }

      val path = s"$pathXls---copy"
      val succes: String = spreadsheet.save(locale, list, path, sheetName, headers, newRow) fold (
        error => throw new IllegalStateException(error),
        success => success)

      succes should be("spreadsheet.service.save.success")

    }

  }

}
