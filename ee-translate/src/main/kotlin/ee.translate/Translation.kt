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
import java.io.File
import java.lang.Double.parseDouble
import java.nio.file.Path
import java.nio.file.Paths
import kotlin.collections.set


private val log = LoggerFactory.getLogger("Trans")

data class Translation(val key: String, var text: String, var index: Int = 0, var bigContext: String = "",
    var contexts: MutableSet<String> = mutableSetOf(),
    var documents: MutableMap<String, Int> = mutableMapOf(),
    var pages: MutableSet<String> = mutableSetOf())

interface TranslationService {
    companion object {
        val NEW_LINE = " @NL@ "
        val REMOVE = "REMOVE"
        val REMOVE_FULL = "REMOVE_FULL"
    }

    fun translate(text: String, context: String = "", document: String = "", page: Int = 0, bigContext: String = "",
        useOriginalAsDefault: Boolean = false): String
}

object TranslationServiceEmptyOrDefault : TranslationService {
    override fun translate(text: String, context: String, document: String, page: Int, bigContext: String,
        useOriginalAsDefault: Boolean): String {
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

    override fun translate(text: String, context: String, document: String, page: Int, bigContext: String,
        useOriginalAsDefault: Boolean): String {
        var translation = translated[text]
        if (translation == null) {
            translation = put(text,
                translationService.translate(text, context, document, page, bigContext, useOriginalAsDefault),
                context, document, page, bigContext)
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
            if (bigContext.isNotEmpty() && translation.bigContext.isEmpty()) {
                translation.bigContext = bigContext
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

    fun put(text: String, translatedText: String, context: String, document: String, page: Int,
        bigContext: String): Translation {
        if(text.trim().isEmpty() || translatedText.isEmpty()) {
            println("empty")
        }
        var ret = translated[text]
        if (ret == null) {
            ret = Translation(text, translatedText, index++, bigContext)
            ret.contexts.add(context)
            ret.documents[document] = 1
            ret.pages.add("1:$page")
            translated[text] = ret
        } else {
            ret.text = translatedText
        }
        return ret
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
                val key = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue
                val translation = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue
                val contexts =
                    row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue.split(separator)
                        .filter { it.trim().isNotEmpty() }.toMutableSet()
                val documents =
                    row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue.split(separator)
                        .filter { it.trim().isNotEmpty() }
                val pages = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue.split(separator)
                    .filter { it.trim().isNotEmpty() }.toMutableSet()
                val bigContext = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).stringCellValue
                val ret = Translation(key, translation, index++, bigContext, contexts, mutableMapOf(), pages)

                documents.forEach {
                    val documentNumber = it.split(documentNumberSeparator)
                    if (documentNumber.size >= 2) {
                        ret.documents[documentNumber[0]] = documentNumber[1].toInt()
                    }
                }
                translated[key] = ret
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
                row.cell(3, it.value.documents.map {
                    "${it.key}$documentNumberSeparator${it.value}"
                }.joinToString(separator))
                row.cell(4, it.value.pages.joinToString(separator))
                row.cell(5, it.value.bigContext)
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
    override fun translate(text: String, context: String, document: String, page: Int, bigContext: String,
        useOriginalAsDefault: Boolean): String {
        return try {
            parseDouble(text)
            text
            //don't translate if number
        } catch (e: Exception) {
            return try {
                translationService.translate(text, context, document, page, bigContext, useOriginalAsDefault)
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

    override fun translate(text: String, context: String, document: String, page: Int, bigContext: String,
        useOriginalAsDefault: Boolean): String {
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

interface FileTranslator<out TextContainer : Any> {
    fun translate(file: File, translationService: TranslationService, targetFile: File, statusUpdater: (String) -> Unit,
        removeTextRun: TextContainer.() -> Boolean = { false })
}


fun <TextContainer : Any> translateFiles(sourceList: List<File>, targetDir: String, dictionaryGlobal: String,
    dictionary: String, languageFrom: String, languageTo: String,
    statusUpdater: (String) -> Unit, removeUnusedFromGlobal: Boolean = false,
    removeTextRun: TextContainer.() -> Boolean = { false },
    translator: FileTranslator<TextContainer>) {
    val target = Paths.get(targetDir)

    val translationServiceRemote = TranslationServiceEmptyOrDefault
    val translationServiceGlobal = TranslationServiceXslx(target.resolve(dictionaryGlobal),
        TranslateServiceNoNeedTranslation(translationServiceRemote))
    var translationService = translationServiceGlobal

    if (dictionary.isNotEmpty()) {
        translationService = TranslationServiceXslx(target.resolve(dictionary), translationServiceGlobal)
    }

    sourceList.forEach { file ->
        translator.translate(file, translationService, target.resolve(file.name).toFile(),
            { statusUpdater("Translate ${file.name}: $it") }, removeTextRun)
    }
    if (translationServiceGlobal != translationService) {
        if (removeUnusedFromGlobal) {
            translationServiceGlobal.removeOtherKeys(translationService.translated.keys)
        }
        translationService.close()
    }

    translationServiceGlobal.close()
}


private val prefix = """(^[ \d:’;.,!%&<>\n\t"/]+)""".toRegex()
private val suffix = """(.+?)([ \d:’;.,!%&<>\n\t"/]+)""".toRegex()

//private val prefix = """(^[ ]+)""".toRegex()
//private val suffix = """(.+?)([ \n\t]+)""".toRegex()

fun translate(raw: String, translationService: TranslationService, context: String, documentName: String,
    pageNumber: Int, bigContext: String): String? {
    var ret: String? = null
    if (raw.trim().isNotEmpty()) {
        var pref = ""
        var suf = ""
        var text = raw
        val prefixAndLastPart = prefix.find(raw)
        if (prefixAndLastPart != null) {
            pref = prefixAndLastPart.groups[1]!!.value
            text = text.removePrefix(pref)
        }

        if (text.isNotEmpty()) {
            val suffixGroups = suffix.matchEntire(text)
            if (suffixGroups != null) {
                text = suffixGroups.groups[1]!!.value
                suf = suffixGroups.groups[2]!!.value
            }

            if (text.isNotEmpty()) {
                val translatedText =
                    translationService.translate(text, context, documentName, pageNumber, bigContext, false)
                log.info("{}={} in '{}'", "$pref$text$suf", translatedText, context)
                if (translatedText.isNotEmpty() && translatedText != text) {
                    try {
                        var translatedFull = "$pref$translatedText$suf"
                        if (translatedText == TranslationService.REMOVE_FULL) {
                            translatedFull = ""
                        } else if (translatedText == TranslationService.REMOVE) {
                            translatedFull = "$pref$suf"
                        }
                        ret = translatedFull
                    } catch (e: Exception) {
                        log.warn("{}", e)
                    }
                }
            }
        }
    }
    return ret
}

