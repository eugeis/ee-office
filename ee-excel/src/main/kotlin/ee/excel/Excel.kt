package ee.excel

import ee.common.ext.exists
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFColor
import org.slf4j.LoggerFactory
import java.io.FileInputStream
import java.io.IOException
import java.net.URL
import java.nio.file.Files
import java.nio.file.Path
import java.text.SimpleDateFormat
import java.util.*
import java.util.regex.Pattern


private val log = LoggerFactory.getLogger(Excel::class.java)
private val dateParser = SimpleDateFormat("DD.MM.YYYY")


class Excel {

    companion object {
        @JvmStatic
        fun open(filePath: Path): Workbook {
            log.info("open '{}'", filePath)
            if (filePath.exists()) {
                return WorkbookFactory.create(FileInputStream(filePath.toFile()))
            } else {
                val ret = HSSFWorkbook()
                ret.createSheet("Worksheet")
                return ret
            }
        }

        @JvmStatic
        fun write(workbook: Workbook, filePath: Path) {
            try {
                Files.newOutputStream(filePath).use {
                    workbook.write(it)
                }
            } catch (e: IOException) {
                throw e
            }
        }

        @JvmStatic
        fun cellIndexToCellName(x: Int, y: Int): String {
            val cellName = dec26(x, 0)
            return cellName + (y + 1)
        }

        @JvmStatic
        private fun dec26(num: Int, first: Int): String {
            return if (num > 25) {
                dec26(num / 26, 1)
            } else {
                ""
            } + ('A' + (num - first) % 26)
        }

        val EMPTY_DATE = Date()
        val EMPTY_URL = URL("http://")

    }
}

operator fun Workbook.get(n: Int): Sheet {
    return this.getSheetAt(n)
}

operator fun Workbook.get(name: String): Sheet {
    return this.getSheet(name)
}

operator fun Sheet.get(n: Int): Row {
    return getRow(n) ?: createRow(n)
}

operator fun Row.get(n: Int): Cell {
    return getCell(n) ?: createCell(n, Cell.CELL_TYPE_BLANK)
}

operator fun Sheet.get(x: Int, y: Int): Cell {
    val row = this[y]
    return row[x]
}

private val ORIGIN = 'A'.toInt()
private val RADIX = 26

// https://github.com/nobeans/gexcelapi/blob/master/src/main/groovy/org/jggug/kobo/gexcelapi/GExcel.groovy
operator fun Sheet.get(cellLabel: String): Cell {
    val p1 = Pattern.compile("([a-zA-Z]+)([0-9]+)");
    val matcher = p1.matcher(cellLabel)
    matcher.find()

    var num = 0
    matcher.group(1).toUpperCase().reversed().forEachIndexed { i, c ->
        val delta = c.toInt() - ORIGIN + 1
        num += delta * Math.pow(RADIX.toDouble(), i.toDouble()).toInt()
    }
    num -= 1
    return this[num, matcher.group(2).toInt() - 1]
}

private fun normalizeNumericString(numeric: Double): String {
    return if (numeric == Math.ceil(numeric)) {
        numeric.toInt().toString()
    } else {
        numeric.toString()
    }
}

fun Cell.toStr(): String {
    when (cellType) {
        Cell.CELL_TYPE_STRING  -> return stringCellValue
        Cell.CELL_TYPE_NUMERIC -> return normalizeNumericString(numericCellValue)
        Cell.CELL_TYPE_BOOLEAN -> return booleanCellValue.toString()
        Cell.CELL_TYPE_BLANK   -> return ""
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_STRING  -> return cellValue.stringValue
                Cell.CELL_TYPE_NUMERIC -> return normalizeNumericString(cellValue.numberValue)
                Cell.CELL_TYPE_BOOLEAN -> return cellValue.booleanValue.toString()
                Cell.CELL_TYPE_BLANK   -> return ""
                else                   -> throw IllegalAccessException("cellはStringに変換できません")
            }

        }
        else                   -> {
            log.warn("Can't parse '$this' to String, return empty.")
            return ""
        }
    }
}

fun Cell.toInt(): Int {
    fun stringToInt(value: String): Int {
        try {
            return value.toDouble().toInt()
        } catch (e: NumberFormatException) {
            log.warn("Can't parse '$this' to Int, return 0")
            return 0
        }
    }

    when (cellType) {
        Cell.CELL_TYPE_STRING  -> return stringToInt(stringCellValue)
        Cell.CELL_TYPE_NUMERIC -> return numericCellValue.toInt()
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_STRING  -> return stringToInt(cellValue.stringValue)
                Cell.CELL_TYPE_NUMERIC -> return cellValue.numberValue.toInt()
                else                   -> throw IllegalAccessException("cellはIntに変換できません")
            }
        }
        else                   -> {
            log.warn("Can't parse '$this' to Int, return 0")
            return 0
        }
    }
}

fun Cell.toDouble(): Double {
    fun stringToDouble(value: String): Double {
        try {
            return value.toDouble()
        } catch (e: NumberFormatException) {
            log.warn("Can't parse '$this' to Double, return 0.0")
            return 0.0
        }
    }

    when (cellType) {
        Cell.CELL_TYPE_STRING  -> return stringToDouble(stringCellValue)
        Cell.CELL_TYPE_NUMERIC -> return numericCellValue.toDouble()
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_STRING  -> return stringToDouble(cellValue.stringValue)
                Cell.CELL_TYPE_NUMERIC -> return cellValue.numberValue.toDouble()
                else                   -> {
                    log.warn("Can't parse '$this' to Double, return 0.0")
                    return 0.0
                }
            }
        }
        else                   -> {
            log.warn("Can't parse '$this' to Double, return 0.0")
            return 0.0
        }
    }
}

fun Cell.toBoolean(): Boolean {
    when (cellType) {
        Cell.CELL_TYPE_BOOLEAN -> return booleanCellValue
        Cell.CELL_TYPE_FORMULA -> {
            val cellValue = getFormulaCellValue(this)
            when (cellValue.cellType) {
                Cell.CELL_TYPE_BOOLEAN -> return cellValue.booleanValue
                else                   -> {
                    log.warn("Can't parse '$this' to Boolean, return false")
                    return false
                }
            }
        }
        else                   -> {
            log.warn("Can't parse '$this' to Boolean, return false")
            return false
        }
    }
}

fun Cell.toDate(): Date {
    try {
        when (cellType) {
            Cell.CELL_TYPE_NUMERIC -> return dateCellValue
            Cell.CELL_TYPE_FORMULA -> {
                val cellValue = getFormulaCellValue(this)
                when (cellValue.cellType) {
                    Cell.CELL_TYPE_NUMERIC -> return dateCellValue
                    else                   -> {
                        log.warn("Can't parse '$this' to Date, return EMPTY")
                        return Excel.EMPTY_DATE
                    }
                }
            }
            Cell.CELL_TYPE_STRING  -> {
                return dateParser.parse(this.stringCellValue)
            }
            else                   -> return Excel.EMPTY_DATE
        }
    } catch (e: Exception) {
        log.warn("Can't parse '$this' to Date, return EMPTY.")
        return Excel.EMPTY_DATE
    }
}

fun Cell.toUrl(): URL {
    try {
        val value = trim()
        if (value.isNotBlank()) {
            return URL(value)
        } else {
            return Excel.EMPTY_URL
        }
    } catch (e: Exception) {
        log.warn("Can't parse '$this' to URL, return null.")
        return Excel.EMPTY_URL
    }
}

fun Cell.trim(): String = toString().trim()
fun Cell.backgroundRgb(): String {
    var ret = ""
    val color = cellStyle.fillBackgroundColorColor
    println(color)
    if (color != null && color is XSSFColor && color.getRGB() != null) {
        val rgb = color.getRGB()
        ret = "rgb(${rgb[0]}, ${rgb[1]}, ${rgb[2]})"
    }
    return ret
}

private fun getFormulaCellValue(cell: Cell): CellValue {
    val workbook = cell.sheet.workbook
    val helper = workbook.creationHelper
    val evaluator = helper.createFormulaEvaluator()
    return evaluator.evaluate(cell)
}

operator fun Sheet.set(cellLabel: String, value: Any) {
    this[cellLabel].setValue(value)
}

operator fun Sheet.set(x: Int, y: Int, value: Any) {
    this[x, y].setValue(value)
}

private fun Cell.setValue(value: Any) {
    when (value) {
        is String  -> setCellValue(value)
        is Int     -> setCellValue(value.toDouble())
        is Double  -> setCellValue(value)
        is Date    -> setCellValue(value)
        is Boolean -> setCellValue(value)
        else       -> throw IllegalArgumentException("Can't set '$value'")
    }
}

fun Row.cell(num: Int, value: String) {
    val cell = getCell(num, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
    cell.setCellType(CellType.STRING)
    cell.setCellValue(value)
}