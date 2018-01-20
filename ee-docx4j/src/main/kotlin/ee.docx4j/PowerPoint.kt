package ee.docx4j

import ee.common.ext.addReturn
import org.docx4j.dml.CTRegularTextRun
import org.docx4j.dml.CTTextCharacterProperties
import org.docx4j.dml.CTTextLineBreak
import org.docx4j.dml.CTTextParagraph
import org.docx4j.openpackaging.packages.PresentationMLPackage
import org.omg.CORBA.Object
import org.slf4j.LoggerFactory
import java.io.File
import java.nio.file.Paths

private val log = LoggerFactory.getLogger("PowerPointDocx4j")


class PowerPoint {
    companion object {
        @JvmStatic
        fun open(fileName: String): PresentationMLPackage {
            return open(Paths.get(fileName).toFile())
        }

        @JvmStatic
        fun open(file: File): PresentationMLPackage {
            return PresentationMLPackage.load(file)
        }
    }
}


class TextRunGroups(val paragraph: CTTextParagraph, val groups: MutableList<TextRunGroup> = mutableListOf()) {
    init {
        var currentTextRunGroup = groups.addReturn(TextRunGroup(paragraph))

        paragraph.egTextRun.forEach {
            if (!currentTextRunGroup.addIfSimilar(it)) {
                currentTextRunGroup = groups.addReturn(TextRunGroup(paragraph))
                currentTextRunGroup.addIfSimilar(it)
            }
        }
    }
}

class TextRunGroup(val paragraph: CTTextParagraph) {
    val textRuns: MutableList<Any> = mutableListOf()
    val textRunsWithOutBaseTextRun: MutableList<Any> = mutableListOf()
    private var baseTextRun: CTRegularTextRun? = null

    fun addIfSimilar(textRun: Any): Boolean {
        var ret = true
        if (textRun.isLineBreak()) {
            textRuns.add(textRun)
            textRunsWithOutBaseTextRun.add(textRun)
        } else if (textRun is CTRegularTextRun) {
            if (baseTextRun == null) {
                baseTextRun = textRun
                textRuns.add(textRun)
            } else if (baseTextRun!!.rPr.isSimilar(textRun.rPr)) {
                textRuns.add(textRun)
                textRunsWithOutBaseTextRun.add(textRun)
            } else {
                ret = false
            }
        }
        return ret
    }

    fun text(): String {
        val out = StringBuilder()
        for (r in textRuns) {
            if (r.isLineBreak()) {
                out.append(" ")
            } else if (r is CTRegularTextRun) {
                out.append(r.t)
            }
        }
        return out.toString()
    }

    fun changeText(text: String) {
        val currentBaseTextRun = baseTextRun
        if (currentBaseTextRun == null) {
            log.info("can't change text, because baseTextRun is null")
            return
        }

        currentBaseTextRun.t = text

        //to check, maybe line break textRun is not needed because calculated bei PowerPoint
        textRunsWithOutBaseTextRun.forEach { it.remove() }
        paragraph.egTextRun.removeAll(textRunsWithOutBaseTextRun)

        textRuns.clear()
        textRuns.add(currentBaseTextRun)
        textRunsWithOutBaseTextRun.clear()
    }

    private fun Any.remove() {
        val items = paragraph.egTextRun
        for (index in 0..items.size - 1) {
            val item = items[index]
            if (this == item) {
                paragraph.egTextRun.remove(this)
                break
            }
        }
    }

    fun removeAllFromParagraph(isToRemove: CTRegularTextRun.() -> Boolean): Boolean {
        val currentBaseTextRun = baseTextRun
        val ret = currentBaseTextRun != null && currentBaseTextRun.isToRemove()
        if (ret) {
            log.info("Remove text runs '{}' from paragraph '{}'", text(), paragraph)
            textRuns.forEach { it.remove() }
            paragraph.egTextRun.removeAll(textRuns)
        }
        return ret
    }

    fun clear() {
        textRuns.clear()
        textRunsWithOutBaseTextRun.clear()
        baseTextRun = null
    }
}

fun Any.isLineBreak(): Boolean {
    return this is CTTextLineBreak
}

fun CTTextCharacterProperties.isSimilar(other: CTTextCharacterProperties): Boolean {
    var ret =
        isB == other.isB && isI == other.isI && isErr == other.isErr && isDirty == other.isDirty && isKumimoji == other.isKumimoji && cs == other.cs //&& fieldType == other.f && fontFamily == this.fontFamily && fontSize == other.fontSize
    if (ret) {
        /*
        if (fontColor is PaintStyle.SolidPaint && other.fontColor is PaintStyle.SolidPaint) {
            val solidColor = (fontColor as PaintStyle.SolidPaint).solidColor
            val otherSolidColor = (other.fontColor as PaintStyle.SolidPaint).solidColor
            ret = solidColor.color == otherSolidColor.color
        } else {
            ret = fontColor == other.fontColor
        }
        */
    }
    return ret
}