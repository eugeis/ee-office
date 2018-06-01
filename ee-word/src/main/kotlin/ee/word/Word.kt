/*
package ee.word

import org.apache.poi.hwpf.HWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.slf4j.LoggerFactory
import java.io.FileInputStream
import java.io.IOException
import java.nio.file.Files
import java.nio.file.Paths
import java.text.SimpleDateFormat

private val log = LoggerFactory.getLogger(Word::class.java)
private val dateParser = SimpleDateFormat("DD.MM.YYYY")


class Word {

    companion object {
        @JvmStatic
        fun isOpen(fileName: String): XWPFDocument {
            return XWPFDocument(FileInputStream(Paths.get(fileName).toFile()))
        }

        @JvmStatic
        fun write(document: XWPFDocument, fileName: String) {
            val outputPath = Paths.get(fileName)
            try {
                Files.newOutputStream(outputPath).use {
                    document.write(it)
                }
            } catch (e: IOException) {
                throw e
            }
        }

        @JvmStatic
        fun openDoc(fileName: String): HWPFDocument {
            return HWPFDocument(FileInputStream(Paths.get(fileName).toFile()))
        }

        @JvmStatic
        fun write(document: HWPFDocument, fileName: String) {
            val outputPath = Paths.get(fileName)
            try {
                Files.newOutputStream(outputPath).use {
                    document.write(it)
                }
            } catch (e: IOException) {
                throw e
            }
        }
    }
}
*/