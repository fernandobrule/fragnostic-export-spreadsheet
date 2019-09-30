package com.fragnostic.export.spreadsheet.api

import java.util.Locale

import org.apache.poi.ss.usermodel.{ Row, Workbook }

/**
 * Created by fernandobrule on 5/19/17.
 */
trait ExportSpreadsheetServiceApi {

  def exportSpreadsheetService: ExportSpreadsheetServiceApi

  trait ExportSpreadsheetServiceApi {

    def spreadsheet[T, S](list: List[T], basePathExport: String, fileName: String, sheetName: String, headers: Array[String], newRow: (Locale, T, Row) => Row): Either[String, String]

    def bytes[T, S](list: List[T], sheetName: String, headers: Array[String], newRow: (Locale, T, Row) => Row): Either[String, Array[Byte]]

    def workbook[T, S](list: List[T], sheetName: String, headers: Array[String], newRow: (Locale, T, Row) => Row): Either[String, Workbook]

  }

}
