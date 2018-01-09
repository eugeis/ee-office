package ee.translate

import com.google.cloud.translate.Translate
import com.google.cloud.translate.Translate.TranslateOption
import com.google.cloud.translate.TranslateOptions
import ee.common.ext.exists
import ee.excel.Excel
import ee.excel.cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.slf4j.LoggerFactory
import java.lang.Double.parseDouble
import java.nio.file.Path
import kotlin.collections.set


private val log = LoggerFactory.getLogger("Trans")

data class Translation(val text: String, var index: Int = 0,
                       var contexts: MutableSet<String> = mutableSetOf(),
                       var documents: MutableMap<String, Int> = mutableMapOf(),
                       var pages: MutableSet<String> = mutableSetOf())

interface TranslationService {
    fun translate(text: String, context: String = "", document: String = "", page: Int = 0, useOriginalAsDefault: Boolean = false): String
}

object TranslationServiceEmptyOrDefault : TranslationService {
    override fun translate(text: String, context: String, document: String, page: Int, useOriginalAsDefault: Boolean): String {
        return if (useOriginalAsDefault) {
            text
        } else {
            ""
        }
    }
}


abstract class AbstractMutableTranslationService(private val translationService: TranslationService,
                                                 val translated: MutableMap<String, Translation> = mutableMapOf()) :
        TranslationService {
    var index: Int = 0

    override fun translate(text: String, context: String, document: String, page: Int, useOriginalAsDefault: Boolean): String {
        var translation = translated[text]
        if (translation == null) {
            translation = Translation(translationService.translate(text, context, document, page, useOriginalAsDefault), index++)
            translation.contexts.add(context)
            translation.documents[document] = 1
            translation.pages.add("1:$page")
            translated[text] = translation
        } else {
            if (context.isNotEmpty()) {
                translation.contexts.add(context)
            }

            if (document.isNotEmpty()) {
                var number = translation.documents[document]
                if (number == null) {
                    number = translation.documents.size + 1
                    translation.documents[document] = number
                }
                translation.pages.add("$number:$page")
            }
        }
        return translation.text;
    }

    fun removeOtherKeys(keys: Set<String>) {
        val size = translated.size
        val toRemove = translated.filterKeys { !keys.contains(it) }
        toRemove.forEach { k, _ ->
            translated.remove(k)
        }
        log.info("removeOtherKeys, original size={}, toRemove={}, current={}", size, toRemove.size, translated.size)
    }
}

class TranslationServiceXslx(private val filePath: Path, translationService: TranslationService,
                             translated: MutableMap<String, Translation> = mutableMapOf()) :
        AbstractMutableTranslationService(translationService, translated) {
    private val separator: String = "≤"
    private val documentNumberSeparator: String = ":"
    private var workbook: Workbook? = null

    init {
        if (filePath.exists()) {
            val currentWorkbook = Excel.open(filePath)
            val sheet = currentWorkbook.getSheetAt(0)
            sheet.forEach { row ->
                val ret = Translation(
                        row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue, index++,
                        row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue.split(separator).filter { it.trim().isNotEmpty() }.toMutableSet(),
                        mutableMapOf(),
                        row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue.split(separator).filter { it.trim().isNotEmpty() }.toMutableSet())

                row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue.split(separator).filter { it.trim().isNotEmpty() }.forEach {
                            val documentNumber = it.split(documentNumberSeparator)
                            if (documentNumber.size >= 2) {
                                ret.documents.put(documentNumber[0], documentNumber[1].toInt())
                            }
                        }
                translated[row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue] = ret
            }
            workbook = currentWorkbook
        }
    }

    fun close() {
        var currentWorkbook: Workbook? = workbook
        if (currentWorkbook == null) {
            currentWorkbook = Excel.open(filePath)
        }

        val sheet = currentWorkbook.getSheetAt(0)
        val indexes = mutableSetOf<Int>()
        translated.forEach {
            if (it.key.trim().isNotEmpty()) {
                var row = sheet.getRow(it.value.index)
                if (row == null) {
                    row = sheet.createRow(it.value.index)
                }
                row.cell(0, it.key)
                row.cell(1, it.value.text)
                row.cell(2, it.value.contexts.joinToString(separator))
                row.cell(3, it.value.documents.map { "${it.key}$documentNumberSeparator${it.value}" }.joinToString(separator))
                row.cell(4, it.value.pages.joinToString(separator))
                indexes.add(it.value.index)
            }
        }

        //remove old indexes
        for (i in 0 until sheet.lastRowNum - 1) { // equivalent of 1 <= i && i <= 10
            if (!indexes.contains(i)) {
                val row = sheet.getRow(i)
                if (row != null) {
                    row.removeAll { true }
                }
            }
        }

        Excel.write(currentWorkbook, filePath)
        currentWorkbook.close()
    }
}

class TranslateServiceNoNeedTranslation(private val translationService: TranslationService) : TranslationService {
    override fun translate(text: String, context: String, document: String, page: Int, useOriginalAsDefault: Boolean): String {
        return try {
            parseDouble(text)
            text
            //don't translate if number
        } catch (e: Exception) {
            return try {
                translationService.translate(text, context, document, page, useOriginalAsDefault)
            } catch (e: Exception) {
                log.warn("can't translate {} because of {}", text, e)
                ""
            }
        }
    }
}

class TranslationServiceByGoogle : TranslationService {
    private var source: TranslateOption
    private var target: TranslateOption
    private val translate: Translate

    constructor(sourceLanguage: String, targetLanguage: String) {
        source = TranslateOption.sourceLanguage(sourceLanguage)
        target = TranslateOption.targetLanguage(targetLanguage)
        translate = TranslateOptions.getDefaultInstance().service
    }

    constructor(sourceLanguage: String, targetLanguage: String, key: String) {
        source = TranslateOption.sourceLanguage(sourceLanguage)
        target = TranslateOption.targetLanguage(targetLanguage)
        translate = TranslateOptions.newBuilder().setApiKey(key).build().service
    }

    override fun translate(text: String, context: String, document: String, page: Int, useOriginalAsDefault: Boolean): String {
        val translation = translate.translate(text, source, target)
        return translation.translatedText
    }

    fun changeSource(sourceLanguage: String) {
        source = TranslateOption.sourceLanguage(sourceLanguage)
    }

    fun changeTarget(targetLanguage: String) {
        target = TranslateOption.sourceLanguage(targetLanguage)
    }
}