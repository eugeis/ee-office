package ee.pptx

import ee.translate.FileTranslator
import ee.translate.TranslationService
import ee.translate.translate
import org.apache.poi.sl.usermodel.PaintStyle
import org.apache.poi.sl.usermodel.TextRun
import org.apache.poi.sl.usermodel.TextShape
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFTextParagraph
import org.slf4j.LoggerFactory
import java.io.File
import java.io.FileOutputStream

private val log = LoggerFactory.getLogger("TranslatePptx")


fun XMLSlideShow.translateTo(translationService: TranslationService, targetFile: File, statusUpdater: (String) -> Unit,
                             removeTextRun: TextRun.() -> Boolean = { false }) {
    val fileName: String = targetFile.nameWithoutExtension
    log.info("translate to {}", fileName)
    slides.forEach { slide ->
        val slideNumber = slide.slideNumber
        log.info("slide: {}", slideNumber)
        statusUpdater("$slideNumber")

        val textShapes = slide.shapes.filterIsInstance(TextShape::class.java)

        val bigContext = textShapes.joinToString("\n") { it.getText() }

        textShapes.forEach { shape ->
            val textParagraphs = shape.getTextParagraphs().filterIsInstance(XSLFTextParagraph::class.java)
            textParagraphs.forEach { paragraph ->

                val rawParagraph = paragraph.text.replace("\n", " ")
                if (rawParagraph.isNotEmpty()) {

                    //check if textRuns can be combined
                    val textRunTranslationGroups = TextRunGroups(paragraph)

                    textRunTranslationGroups.groups.filterNot { it.removeAllFromParagraph(removeTextRun) }
                        .forEach { group ->
                            val translated =
                                translate(group.text(), translationService, rawParagraph, fileName, slideNumber,
                                    bigContext)
                            if (translated != null) {
                                group.changeText(translated)
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

class PptxFileTranslator : FileTranslator<TextRun> {
    override fun translate(file: File, translationService: TranslationService, targetFile: File,
                           statusUpdater: (String) -> Unit, removeTextRun: TextRun.() -> Boolean) {
        try {
            val slideShow = PowerPoint.open(file)
            slideShow.translateTo(translationService, targetFile, { statusUpdater("Translate ${file.name}: $it") },
                removeTextRun)
        } catch (e: Exception) {
            log.warn("Can't translate '{}' because of '{}'", file, e)
        }
    }
}

fun TextRun.isColor(red: Int = 0, green: Int = 0, blue: Int = 0): Boolean {
    var ret = false
    if (fontColor is PaintStyle.SolidPaint) {
        val color = (fontColor as PaintStyle.SolidPaint).solidColor.color
        ret = color.red == red && color.green == green && color.blue == blue
    }
    return ret
}
