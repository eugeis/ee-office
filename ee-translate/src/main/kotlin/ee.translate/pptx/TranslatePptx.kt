package ee.translate.pptx

import ee.common.ext.addReturn
import ee.pptx.PowerPoint
import ee.pptx.TextRunGroup
import ee.translate.TranslateServiceNoNeedTranslation
import ee.translate.TranslationService
import ee.translate.TranslationServiceEmptyOrDefault
import ee.translate.TranslationServiceXslx
import org.apache.poi.sl.usermodel.PaintStyle
import org.apache.poi.sl.usermodel.TextRun
import org.apache.poi.sl.usermodel.TextShape
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFTextParagraph
import org.slf4j.LoggerFactory
import java.io.File
import java.io.FileOutputStream
import java.nio.file.Paths

private val log = LoggerFactory.getLogger("TranslatePptx")

private val REMOVE = "REMOVE"
private val REMOVE_FULL = "REMOVE_FULL"

private val prefix = """(^[ \d:’“;.,«»?!%&<>\n\t\-—+'…\[\]"/°№„]+)""".toRegex()
private val suffix = """(.+?)([ \d:’“;.,«»?!%&<>\n\t\-—+'…\[\]"/°№„]+)""".toRegex()

fun XMLSlideShow.translateTo(translationService: TranslationService, targetFile: File,
                             statusUpdater: (String) -> Unit,
                             removeTextRun: TextRun.() -> Boolean = { false }) {
    val fileName: String = targetFile.nameWithoutExtension
    log.info("translate to {}", fileName)
    slides.forEach { slide ->
        val slideNumber = slide.slideNumber
        log.info("slide: {}", slideNumber)
        statusUpdater("$slideNumber")

        slide.shapes.filterIsInstance(TextShape::class.java).forEach { shape ->
                    shape.getTextParagraphs().filterIsInstance(XSLFTextParagraph::class.java).forEach { paragraph ->

                                val rawParagraph = paragraph.text.replace("\n", " ")
                                if (rawParagraph.isNotEmpty()) {

                                    //check if textRuns can be combined
                                    val textRunTranslationGroups = mutableListOf<TextRunGroup>()
                                    var currentTextRunGroup = textRunTranslationGroups.addReturn(TextRunGroup(paragraph))

                                    paragraph.textRuns.forEach {
                                        if (!currentTextRunGroup.addIfSimilar(it)) {
                                            currentTextRunGroup = textRunTranslationGroups.addReturn(TextRunGroup(paragraph))
                                            currentTextRunGroup.addIfSimilar(it)
                                        }
                                    }

                                    textRunTranslationGroups.filterNot { it.removeAllFromParagraph(removeTextRun) }.forEach { group ->
                                                val raw = group.text()
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
                                                            var translatedText = translationService.translate(text,
                                                                    rawParagraph, fileName, slideNumber, false)
                                                            log.info("{}={} in '{}'", "$pref$text$suf", translatedText, rawParagraph)
                                                            if (translatedText.isNotEmpty() && translatedText != text) {
                                                                try {
                                                                    var translatedFull = "$pref$translatedText$suf"
                                                                    if (translatedText == REMOVE_FULL) {
                                                                        translatedFull = ""
                                                                    } else if (translatedText == REMOVE) {
                                                                        translatedFull = "$pref$suf"
                                                                    }
                                                                    group.changeText(translatedFull)
                                                                } catch (e: Exception) {
                                                                    log.warn("{}", e)
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                }
                            }
                }
    }

    val fileOut = FileOutputStream(targetFile)
    write(fileOut)
    fileOut.close()
}

fun translatePowerPoints(sourceDir: String, targetDir: String, dictionaryGlobal: String, dictionary: String,
                         languageFrom: String, languageTo: String,
                         statusUpdater: (String) -> Unit, removeTextRun: TextRun.() -> Boolean = { false }) {
    val target = Paths.get(targetDir)

    val translationServiceRemote = TranslationServiceEmptyOrDefault
    val translationServiceGlobal = TranslationServiceXslx(target.resolve(dictionaryGlobal),
            TranslateServiceNoNeedTranslation(translationServiceRemote))
    val translationService = TranslationServiceXslx(target.resolve(dictionary), translationServiceGlobal)

    Paths.get(sourceDir).toFile().listFiles { file -> file.name.endsWith(".pptx", true) }.forEach { file ->
                try {
                    val slideShow = PowerPoint.open(file)
                    slideShow.translateTo(translationService, target.resolve(file.name).toFile(),
                            { statusUpdater("Translate ${file.name}: $it") }, removeTextRun)
                } catch (e: Exception) {
                    log.warn("Can't translate '{}' because of '{}'", file, e)
                }
            }
    //translationServiceGlobal.removeOtherKeys(translationService.translated.keys)

    translationService.close()
    translationServiceGlobal.close()
}

fun TextRun.isColor(red: Double = 0.0, green: Double = 0.0, blue: Double = 0.0): Boolean {
    var ret = false
    if (fontColor is PaintStyle.SolidPaint) {
        val color = (fontColor as PaintStyle.SolidPaint).solidColor.color
        ret = color.red == red.toInt() && color.green == green.toInt() && color.blue == blue.toInt()
    }
    return ret
}