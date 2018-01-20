package ee.docx4j

import ee.common.ext.addReturn
import org.apache.xpath.operations.Bool
import org.docx4j.dml.*
import org.docx4j.openpackaging.packages.PresentationMLPackage
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


class TextGroups(val paragraph: CTTextParagraph, val items: MutableList<TextGroup> = mutableListOf(),
    private var text: String? = null) {
    init {
        if (paragraph.egTextRun.isNotEmpty()) {
            var currentTextRunGroup = items.addReturn(TextGroup(paragraph))

            paragraph.egTextRun.forEach {
                if (!currentTextRunGroup.addIfSimilar(it)) {
                    currentTextRunGroup = items.addReturn(TextGroup(paragraph))
                    currentTextRunGroup.addIfSimilar(it)
                }
            }
        }
    }

    fun text(): String {
        if (text == null) {
            text = items.joinToString(" ") { it.text() }
        }
        return text!!
    }

    fun clear() {
        text = null
    }
}

class TextGroup(val paragraph: CTTextParagraph, private var text: String? = null) {
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
        if (text == null) {
            val out = StringBuilder()
            for (r in textRuns) {
                if (r.isLineBreak()) {
                    out.append(" ")
                } else if (r is CTRegularTextRun) {
                    out.append(r.t)
                }
            }
            text = out.toString()
        }
        return text!!
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
        text = null
    }
}

fun Any.isLineBreak(): Boolean {
    return this is CTTextLineBreak
}

fun CTTextCharacterProperties.isSimilar(other: CTTextCharacterProperties): Boolean {
    var ret = this.ln == other.ln && this.noFill == other.noFill && this.solidFill.isSimilar(other.solidFill) &&
            this.gradFill == other.gradFill && this.blipFill == other.blipFill && this.pattFill == other.pattFill &&
            this.grpFill == other.grpFill && this.effectLst == other.effectLst && this.effectDag == other.effectDag &&
            this.highlight == other.highlight && this.uLnTx == other.uLnTx && this.uLn == other.uLn &&
            this.uFillTx == other.uFillTx && this.uFill == other.uFill && this.latin == other.latin &&
            this.ea == other.ea && this.cs == other.cs && this.sym == other.sym &&
            this.hlinkClick == other.hlinkClick && this.hlinkMouseOver == other.hlinkMouseOver &&
            this.extLst == other.extLst && this.isKumimoji == other.isKumimoji &&
            this.altLang == other.altLang && this.sz == other.sz && this.isB == other.isB && this.isI == other.isI &&
            this.u == other.u && this.strike == other.strike && this.kern == other.kern && this.cap == other.cap &&
            this.spc == other.spc && this.isNormalizeH == other.isNormalizeH && this.baseline == other.baseline &&
            this.isNoProof == other.isNoProof && this.isDirty == other.isDirty &&
            this.smtId == other.smtId && this.bmk == other.bmk
    //&& this.isErr == other.isErr && this.isSmtClean == other.isSmtClean && this.lang == other.lang
    return ret
}

fun CTSolidColorFillProperties?.isSimilar(other: CTSolidColorFillProperties?): Boolean {
    val ret = this == other || this?.srgbClr?.`val`==other?.srgbClr?.`val`
    return ret
}