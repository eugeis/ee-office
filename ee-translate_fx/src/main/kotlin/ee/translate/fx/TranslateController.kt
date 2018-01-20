package ee.translate.fx

import ee.pptx.PptxFileTranslator
import ee.pptx.collectPowerPointFiles
import ee.pptx.isColor
import ee.translate.translateFiles
import javafx.application.Platform
import org.apache.poi.sl.usermodel.TextRun
import tornadofx.*

class TranslateController : Controller() {
    val dashboard: Dashboard by inject()

    fun init() {
        showDashboard()
    }

    fun showDashboard() {
        if (FX.primaryStage.scene.root != dashboard.root) {
            FX.primaryStage.scene.root = dashboard.root
            FX.primaryStage.sizeToScene()
            FX.primaryStage.centerOnScreen()
            FX.primaryStage.title = dashboard.title
        }

        Platform.runLater {
            with(config) {
                if (containsKey(SOURCE_DIR)) {
                    dashboard.sourceDirOrFiles.text = string(SOURCE_DIR)
                }
                if (containsKey(TARGET_DIR)) {
                    dashboard.targetDir.text = string(TARGET_DIR)
                }
                if (containsKey(DICTIONARY_GLOBAL)) {
                    dashboard.dictionaryGlobal.text = string(DICTIONARY_GLOBAL)
                }
                if (containsKey(DICTIONARY)) {
                    dashboard.dictionary.text = string(DICTIONARY)
                }
                if (containsKey(LANGUAGE_FROM)) {
                    dashboard.languageFrom.text = string(LANGUAGE_FROM)
                }
                if (containsKey(LANGUAGE_TO)) {
                    dashboard.languageTo.text = string(LANGUAGE_TO)
                }
            }
        }
    }

    fun translate() {
        var removeTextRun: TextRun.() -> Boolean = { false }
        if (dashboard.removeByColor.isSelected) {
            val color = dashboard.colorToRemove.value
            val red = (color.red * 255).toInt()
            val green = (color.green * 255).toInt()
            val blue = (color.blue * 255).toInt()
            removeTextRun = { isColor(red, green, blue) }
        }

        val files = collectPowerPointFiles(dashboard.sourceDirOrFiles.text, dashboard.delimiter)

        val fileTranslator = PptxFileTranslator()

        translateFiles(files, dashboard.targetDir.text, dashboard.dictionaryGlobal.text,
            dashboard.dictionary.text, dashboard.languageFrom.text, dashboard.languageTo.text, dashboard.statusUpdater,
            dashboard.removeUnusedFromGlobal.isSelected, removeTextRun, fileTranslator)
    }

    fun storeSettings() {
        with(config) {
            set(SOURCE_DIR to dashboard.sourceDirOrFiles.text)
            set(TARGET_DIR to dashboard.targetDir.text)
            set(DICTIONARY_GLOBAL to dashboard.dictionaryGlobal.text)
            set(DICTIONARY to dashboard.dictionary.text)
            set(LANGUAGE_FROM, dashboard.languageFrom.text)
            set(LANGUAGE_TO, dashboard.languageTo.text)
            save()
        }
    }

    fun clearSettings() {
        with(config) {
            remove(SOURCE_DIR)
            remove(TARGET_DIR)
            remove(DICTIONARY_GLOBAL)
            remove(DICTIONARY)
            remove(LANGUAGE_FROM)
            remove(LANGUAGE_TO)
            save()
        }
    }

    fun exit() {
        storeSettings()
        Platform.exit()
    }

    companion object {
        val SOURCE_DIR = "SOURCE_DIR"
        val TARGET_DIR = "TARGET_DIR"
        val DICTIONARY_GLOBAL = "DICTIONARY_GLOBAL"
        val DICTIONARY = "DICTIONARY"
        val LANGUAGE_FROM = "LANGUAGE_FROM"
        val LANGUAGE_TO = "LANGUAGE_TO"
    }
}