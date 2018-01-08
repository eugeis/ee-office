package ee.translate.pptx

import ee.common.ext.addReturn
import ee.common.ext.withFileNameSuffix
import ee.pptx.PowerPoint
import ee.pptx.TextRunGroup
import ee.translate.TranslateServiceNoNeedTranslation
import ee.translate.TranslationService
import ee.translate.TranslationServiceEmpty
import ee.translate.TranslationServiceXslx
import org.apache.poi.sl.usermodel.TextRun
import org.apache.poi.sl.usermodel.TextShape
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFTextParagraph
import org.slf4j.LoggerFactory
import java.io.File
import java.io.FileOutputStream
import java.nio.file.Paths

private val log = LoggerFactory.getLogger("TranslatePptx")

private val prefix = """(^[ \d:’“;.,«»?!%&<>\n\t)(\-—+'…\[\]"/°№„]+)""".toRegex()
private val suffix = """(.+?)([ \d:’“;.,«»?!%&<>\n\t)(\-—+'…\[\]"/°№„]+)""".toRegex()

fun XMLSlideShow.translateTo(translationService: TranslationService, targetFile: File,
                             removeTextRun: TextRun.() -> Boolean = { false }) {
    val fileName: String = targetFile.nameWithoutExtension
    log.info("translate to {}", fileName)
    slides.forEach { slide ->
        val slideNumber = slide.slideNumber
        log.info("slide: {}", slideNumber)

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
                                    val translated = translationService.translate(text, rawParagraph, fileName, slideNumber)
                                    log.info("{}={} in '{}'", text, translated, rawParagraph)
                                    if (translated.isNotEmpty() && translated != text) {
                                        try {
                                            group.changeText("$pref$translated$suf")
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

fun translatePowerpoints(sourceDir: String, targetDir: String, from: String, to: String, googleKey: String,
                         ignoreTextRun: TextRun.() -> Boolean = { false }) {
    val target = Paths.get(targetDir)

    //val translationServiceRemote = TranslationServiceByGoogle(from, to, googleKey)
    val translationServiceRemote = TranslationServiceEmpty
    val translationServiceGlobal = TranslationServiceXslx(target.resolve("dictionary_global.xlsx"),
            TranslateServiceNoNeedTranslation(translationServiceRemote))
    val translationService = TranslationServiceXslx(target.resolve("dictionary.xlsx"), translationServiceGlobal)
    //val translationService = translationServiceGlobal

    Paths.get(sourceDir).toFile().listFiles { file -> file.name.endsWith(".pptx", true) }.forEach { file ->
        try {
            val slideShow = PowerPoint.open(file)
            val targetFileName = file.name.withFileNameSuffix("_" + to)
            slideShow.translateTo(translationService, target.resolve(targetFileName).toFile(), ignoreTextRun)
        } catch (e: Exception) {
            log.warn("Can't translate '{}' because of '{}'", file, e)
        }
    }
    translationServiceGlobal.removeOtherKeys(translationService.translated.keys)

    translationService.close()
    translationServiceGlobal.close()
}