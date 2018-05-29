package ee.docx4j

import ee.translate.FileTranslator
import ee.translate.TranslationService
import ee.translate.translate
import org.docx4j.dml.CTRegularTextRun
import org.docx4j.dml.CTTextCharacterProperties
import org.docx4j.dml.diagram.CTRelIds
import org.docx4j.openpackaging.packages.PresentationMLPackage
import org.docx4j.openpackaging.parts.DrawingML.DiagramDataPart
import org.docx4j.openpackaging.parts.Part
import org.pptx4j.pml.CTGraphicalObjectFrame
import org.pptx4j.pml.Shape
import org.slf4j.LoggerFactory
import java.io.File
import javax.xml.bind.JAXBElement

private val log = LoggerFactory.getLogger("TranslatePptx")

fun PresentationMLPackage.translateTo(translationService: TranslationService, targetFile: File,
    statusUpdater: (String) -> Unit,
    removeTextRun: CTRegularTextRun.() -> Boolean = { false }) {
    val documentName: String = targetFile.nameWithoutExtension
    log.info("translate to {}", documentName)

    for (i in 0 until mainPresentationPart.slideCount - 1) {
        val pageNumber = i + 1
        val slide = mainPresentationPart.getSlide(i)
        statusUpdater("slide $pageNumber")

        slide.contents.cSld.spTree.spOrGrpSpOrGraphicFrame.forEach { shape ->
            when (shape) {
                is CTGraphicalObjectFrame -> {
                    val data = shape.graphic.graphicData
                    val relIds =
                        (data.any?.find { it is JAXBElement<*> && it.value is CTRelIds } as JAXBElement<*>?)?.value as CTRelIds?
                    if (relIds != null) {
                        val part = slide.relationships.getPart(relIds.dm)
                        part.translate(translationService, documentName, pageNumber, removeTextRun)
                    }
                }
                is Shape -> {
                    val textGroupsList = mutableListOf<TextGroups>()
                    shape.txBody?.p?.filter { it.egTextRun.isNotEmpty() }?.forEach { textGroupsList.add(TextGroups(it)) }
                    textGroupsList.translate(translationService, documentName, pageNumber, removeTextRun)
                }
                else -> {
                    log.info("{}", shape)
                }
            }
        }
    }
    this.save(targetFile)
}

private fun Part.translate(translationService: TranslationService, documentName: String, pageNumber: Int,
    removeTextRun: CTRegularTextRun.() -> Boolean) {
    when (this) {
        is DiagramDataPart -> {
            val textGroupsList = mutableListOf<TextGroups>()
            contents.ptLst.pt.filter { it?.t?.p != null && it.t.p.isNotEmpty() }.forEach { pt ->
                pt.t.p.filter { it.egTextRun.isNotEmpty() }.forEach { textGroupsList.add(TextGroups(it)) }
            }
            textGroupsList.translate(translationService, documentName, pageNumber, removeTextRun)
        }
        else -> log.info("Else: {}", this)
    }
}

private fun List<TextGroups>.translate(translationService: TranslationService, documentName: String, pageNumber: Int,
    removeTextRun: CTRegularTextRun.() -> Boolean) {
    if (isNotEmpty()) {
        val bigContext = joinToString("\n") { it.text() }
        forEach { textGroups ->
            val context = textGroups.text()
            textGroups.items.filter { it.textRuns.isNotEmpty() && !it.removeAllFromParagraph(removeTextRun) }
                .forEach { group ->
                    val translated =
                        translate(group.text(), translationService, context, documentName, pageNumber, bigContext)
                    if (translated != null) {
                        group.changeText(translated)
                    }
                }
        }
    }
}


class Docx4jPptxFileTranslator : FileTranslator<CTRegularTextRun> {
    override fun translate(file: File, translationService: TranslationService, targetFile: File,
        statusUpdater: (String) -> Unit, removeTextRun: CTRegularTextRun.() -> Boolean) {
        try {
            val slideShow = PowerPoint.open(file)
            slideShow.translateTo(translationService, targetFile, { statusUpdater("Translate ${file.name}: $it") },
                removeTextRun)
        } catch (e: Exception) {
            log.warn("Can't translate '{}' because of '{}'", file, e)
        }
    }
}

fun CTTextCharacterProperties.isColor(rgbColorValue: String): Boolean {
    return rgbColorValue == solidFill?.srgbClr?.`val`
}