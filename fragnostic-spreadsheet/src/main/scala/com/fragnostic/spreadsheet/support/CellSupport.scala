package com.fragnostic.spreadsheet.support

import org.apache.poi.ss.usermodel.Row

/**
 * Created by fernandobrule on 5/28/17.
 */
trait CellSupport {

  def createCell(row: Row, cellNum: Int, cellValue: String, enabled: Boolean): Int =
    if (enabled) {
      row.createCell(cellNum).setCellValue(cellValue)
      cellNum + 1
    } else {
      cellNum
    }

  def createCell(row: Row, cellNum: Int, cellValue: Double, enabled: Boolean): Int =
    if (enabled) {
      row.createCell(cellNum).setCellValue(cellValue)
      cellNum + 1
    } else {
      cellNum
    }

}
