package com.fragnostic.export.excel.api

import java.util.Locale

import org.apache.poi.ss.usermodel.{ Row, Workbook }

/**
 * Created by fernandobrule on 5/19/17.
 */
trait ExportExcelServiceApi {

  def exportExcelService: ExportExcelServiceApi

  trait ExportExcelServiceApi {

    def export[T, S](list: List[T], fileName: String, sheetName: String, headers: Array[String], enables: S, newRow: (Locale, T, Row, S) => Row): Either[String, String]

    def exportBytes[T, S](list: List[T], sheetName: String, headers: Array[String], enables: S, newRow: (Locale, T, Row, S) => Row): Either[String, Array[Byte]]

    def exportWb[T, S](list: List[T], sheetName: String, headers: Array[String], enables: S, newRow: (Locale, T, Row, S) => Row): Either[String, Workbook]

  }

}
