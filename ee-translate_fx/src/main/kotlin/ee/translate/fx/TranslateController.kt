package ee.translate.fx

import ee.translate.pptx.isColor
import ee.translate.pptx.translatePowerPoints
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
                    dashboard.sourceDir.text = string(SOURCE_DIR)
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
            removeTextRun = {
                val color = dashboard.colorToRemove.value
                isColor(color.red, color.green, color.blue)
            }
        }

        translatePowerPoints(dashboard.sourceDir.text, dashboard.targetDir.text,
                dashboard.dictionaryGlobal.text, dashboard.dictionary.text,
                dashboard.languageFrom.text, dashboard.languageTo.text,
                dashboard.statusUpdater, removeTextRun)
    }

    fun storeSettings() {
        with(config) {
            set(SOURCE_DIR to dashboard.sourceDir.text)
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