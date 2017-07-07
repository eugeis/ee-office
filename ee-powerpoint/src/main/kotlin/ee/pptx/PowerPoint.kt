package ee.pptx

import ee.common.ext.ext
import ee.common.ext.ifElse
import ee.common.ext.letTraceExc
import ee.common.ext.toKey
import ee.slides.*
import org.apache.poi.sl.usermodel.PaintStyle
import org.apache.poi.xslf.usermodel.*
import java.awt.geom.Rectangle2D
import java.io.File
import java.io.FileInputStream
import java.nio.file.Path
import java.nio.file.Paths
import java.util.*

private var colors = hashMapOf<String, Color>()
private var anchors = hashMapOf<String, Rectangle>()
private var fonts = hashMapOf<String, Font>()

class PowerPoint {
    companion object {
        @JvmStatic
        fun open(fileName: String): XMLSlideShow {
            return open(Paths.get(fileName).toFile())
        }

        @JvmStatic
        fun open(file: File): XMLSlideShow {
            return XMLSlideShow(FileInputStream(file))
        }

        fun parseFiles(path: Path): List<Presentation> {
            val presentations: List<Presentation> = path.toFile().walkTopDown().filter(File::isPresentation).
                    mapNotNull(File::toPresentation).toList()
            return presentations
        }

        fun parseFilesAsTopics(path: Path, name: String): Presentation {
            return path.toFile().walkTopDown().filter(File::isPresentation).sorted().toPresentation(name)
        }
    }
}

fun XMLSlideShow.toPresentation(name: String): Presentation {
    resetCaches()

    val ret = Presentation(name = name, topics = arrayListOf(Topic(name,
            slides = slides.mapNotNull(XSLFSlide::toSlide).toMutableList())))

    ret.assignCaches()

    return ret
}

private fun resetCaches() {
    colors = hashMapOf()
    anchors = hashMapOf()
    fonts = hashMapOf()
}

private fun Presentation.assignCaches() {
    anchors = ArrayList<Rectangle>(ee.pptx.anchors.values)
    fonts = ArrayList<Font>(ee.pptx.fonts.values)
    colors = ArrayList<Color>(ee.pptx.colors.values)
}

fun XMLSlideShow.toTopic(name: String): Topic = Topic(name = name,
        slides = slides.mapNotNull(XSLFSlide::toSlide).toMutableList())

fun XSLFSlide.toSlide(): Slide? = letTraceExc {
    //often is title same as paragraphs of first shape, so check it out and filter it out
    var title = title.orEmpty()
    val shapes = shapes.mapNotNull(XSLFShape::toShape).toMutableList()
    if (title.isNotEmpty()) {
        val titleShape = shapes.find { it is TextShape && "TITLE".equals(it.textType, ignoreCase = true) }
        if (titleShape != null) {
            shapes.remove(titleShape)
        }
    }

    //val masterName = masterSheet?.name.orEmpty().toKey()
    val masterType = masterSheet?.type?.name?.toKey()

    Slide(title = title, masterType = masterType.orEmpty(), shapes = shapes, notes = notes?.toNotes().orEmpty())
}

fun XSLFNotes.toNotes(): Notes = Notes(shapes = shapes.mapNotNull(XSLFShape::toShape).toMutableList())

fun XSLFComments.toComments(): MutableList<String> = ctCommentsList.cmList.mapNotNull { it.letTraceExc { it.text } }.toMutableList()

fun Rectangle2D.toAnchor(): Rectangle? = letTraceExc {
    val ret = "${height}_${width}__${x}_${y}"
    anchors.getOrPut(ret, {
        Rectangle(name = ret, height = height.toInt(), width = width.toInt(), x = x.toInt(), y = y.toInt())
    })
}

fun XSLFTextParagraph.toParagraphType(): ParagraphType {
    if (autoNumberingScheme != null) {
        return ParagraphType.NUMBERED
    } else if (isBullet) {
        return ParagraphType.BULLET
    } else {
        return ParagraphType.DEFAULT
    }
}

fun PaintStyle.toColor(): Color {
    if (this is PaintStyle.SolidPaint) {
        val color = this.solidColor.color
        val ret = "${color.red}_${color.green}_${color.blue}"
        return colors.getOrPut(ret, {
            Color(red = color.red, blue = color.blue, green = color.green, alpha = color.alpha)
        })
    } else {
        return Color.EMPTY
    }
}

fun XSLFTextRun.toFont(): Font? = letTraceExc {
    val ret = "${fontFamily.toKey()}${isItalic.ifElse("_italic", "")}${isBold.ifElse("_bold", "")}${isUnderlined.ifElse("_underlined", "")}"
    fonts.getOrPut(ret, {
        Font(name = ret, family = fontFamily,
                italic = isItalic, underlined = isUnderlined, bold = isBold)
    })
}

fun XSLFTextRun.toTextRun(): TextRun? = letTraceExc {
    TextRun(text = rawText, font = toFont()?.name.orEmpty(), color = fontColor.toColor().name, cap = textCap.name.toTextCap())
}

fun XSLFTextParagraph.toParagraph(): Paragraph? = letTraceExc {
    Paragraph(type = toParagraphType(),
            textAlign = textAlign?.name.toTextAlign(),
            fontAlign = fontAlign?.name.toFontAlign(), textRuns = textRuns.mapNotNull { it.toTextRun() }.toMutableList())
}

fun XSLFShape.toShape(): Shape? = letTraceExc {
    val name = shapeName.orEmpty()
    if (this is XSLFTextShape) {
        TextShape(name = name, textType = this.textType?.name.orEmpty(), type = ShapeType.TEXT,
                anchor = anchor?.toAnchor().orEmpty().name,
                paragraphs = textParagraphs.mapNotNull {
                    it.letTraceExc { it.toParagraph() }
                }.toMutableList())
    } else if (this is XSLFGroupShape) {
        GroupShape(name = name, type = ShapeType.GROUP, shapes = shapes.mapNotNull(XSLFShape::toShape).toMutableList())
    } else if (this is XSLFPictureShape) {
        if (isExternalLinkedPicture) {
            PictureShape(name = name, link = true, linkUri = pictureLink.toString(), type = ShapeType.PICTURE)
        } else {
            PictureShape(name = name, link = false, data = pictureData.data, type = ShapeType.PICTURE)
        }
    } else if (this is XSLFGraphicFrame) {
        println("Shape type supported yet $this")
        GraphicShape(name = name, type = ShapeType.GRAPHIC)
    } else if (this is XSLFTable) {
        println("Shape type supported yet $this")
        GraphicShape(name = name, type = ShapeType.TABLE)
    } else {
        println("Shape type supported yet $this")
        null
    }
}

fun File.isPresentation() = !(name.startsWith(".") || name.startsWith("~")) && ext().equals("pptx")

fun File.toTopic(): Topic? = letTraceExc { PowerPoint.open(this).toTopic(nameWithoutExtension) }

fun Sequence<File>.toPresentation(name: String): Presentation {
    resetCaches()
    val ret = Presentation(name, topics = mapNotNull { it.toTopic() }.toMutableList())
    ret.assignCaches()
    return ret
}

fun File.toPresentation(): Presentation? = letTraceExc { PowerPoint.open(this).toPresentation(nameWithoutExtension) }