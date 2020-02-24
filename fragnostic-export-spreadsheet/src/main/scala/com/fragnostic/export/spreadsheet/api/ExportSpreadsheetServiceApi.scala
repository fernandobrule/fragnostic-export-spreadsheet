package com.fragnostic.export.spreadsheet.api

import java.util.Locale

import org.apache.poi.ss.usermodel.{ Row, Workbook }

/**
 * Created by fernandobrule on 5/19/17.
 */
trait ExportSpreadsheetServiceApi {

  def exportSpreadsheetService: ExportSpreadsheetServiceApi

  trait ExportSpreadsheetServiceApi {

    def getBytes[T, S](locale: Locale, list: List[T], sheetName: String, headers: Array[String], newRow: (Locale, T, Row) => Row): Either[String, Array[Byte]]

    def getBytes(workbook: Workbook): Either[String, Array[Byte]]

    def getWorkbook[T, S](locale: Locale, list: List[T], sheetName: String, headers: Array[String], newRow: (Locale, T, Row) => Row): Either[String, Workbook]

    def getWorkbook(path: String): Either[String, Workbook]

    def getWorkbook(bytes: Array[Byte]): Either[String, Workbook]

    def save[T, S](locale: Locale, list: List[T], basePathExport: String, fileName: String, sheetName: String, headers: Array[String], newRow: (Locale, T, Row) => Row): Either[String, String]

    def save(bytes: Array[Byte], path: String): Either[String, String]

    def save(workbook: Workbook, path: String): Either[String, String]

  }

}
