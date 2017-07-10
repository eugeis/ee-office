package ee.word

import ee.common.ext.fileExt
import ee.word.Word.Companion.openDoc
import org.apache.poi.hwpf.HWPFDocument
import org.slf4j.LoggerFactory
import java.nio.file.Paths

data class Content(var name: String, val points: MutableList<String>)

private val log = LoggerFactory.getLogger("Wordtest")

fun main(args: Array<String>) {
    val path = Paths.get("D:\\TC_CACHE\\CG-latest\\doc\\clearing\\2017-01-19 ClearingReports")
    val contents = arrayListOf<Content>()
    path.toFile().listFiles { file -> file.name.fileExt().equals("doc", true) }.forEach { file ->
        log.info("Extract $file")
        var doc: HWPFDocument? = null
        try {
            doc = openDoc(file.absolutePath)
            val content = doc.extractChapter()
            content.name = file.name
            contents.add(content)
        } catch (e: Exception) {
            contents.add(Content(file.name, arrayListOf("ERROR AT FILE PARSING $e")))
        } finally {
            if (doc != null) {
                doc.close()
            }
        }
    }
    contents.forEach { doc ->
        log.info(doc.name)
        doc.points.forEach { point ->
            log.info("\t$point")
        }

    }
    log.info("$contents")
}

private fun HWPFDocument.extractChapter(): Content {
    val startExpr = ".*(\\d+?\\.\\d+?\\.).*Obligations.*".toRegex(RegexOption.DOT_MATCHES_ALL)
    val endExpr = ".*(\\d+?\\.).*Notes.*".toRegex(RegexOption.DOT_MATCHES_ALL)

    val points = arrayListOf<String>()
    var started = false
    for (i in 0 until range.numParagraphs()) {
        val item = range.getParagraph(i)
        val text = item.text()
        if (text.contains("Obligations")) {
            log.debug(text)
        }
        if (!started && startExpr.matches(text)) {
            started = true
        } else if (started) {
            if (!endExpr.matches(text)) {
                points.add(text)
            } else {
                break
            }
        }
    }
    return Content("", points)
}
