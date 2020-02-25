package com.fragnostic.spreadsheet

import java.util.Locale

import com.fragnostic.spreadsheet.CakeServiceSpreadsheet.spreadsheet
import com.fragnostic.spreadsheet.support.{ AbstractTest, SomeRow }
import org.apache.poi.ss.usermodel.Row

class GetBytesFromDataTest extends AbstractTest {

  describe("Get Bytes From Data Test") {

    it("Can Get Get Bytes From Data") {

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

      val bytes: Array[Byte] = spreadsheet.getBytes(locale, list, sheetName, headers, newRow) fold (
        error => throw new IllegalStateException(error),
        bytes => bytes)

      val ruta = s"$basePath/workbook-$randomInt.xls"
      fileExists(ruta) should be(false)

      spreadsheet.save(bytes, ruta)

      fileExists(ruta) should be(true)

    }

  }

}
