package ee.translate.fx

import javafx.application.Platform
import javafx.beans.property.SimpleStringProperty
import javafx.beans.property.StringProperty
import javafx.geometry.Insets
import javafx.geometry.Pos
import javafx.scene.control.CheckBox
import javafx.scene.control.ColorPicker
import javafx.scene.control.ContentDisplay
import javafx.scene.control.TextField
import javafx.scene.layout.BorderPane
import javafx.scene.layout.Priority
import javafx.scene.paint.Color
import javafx.stage.DirectoryChooser
import javafx.stage.FileChooser
import tornadofx.*

class Dashboard : View() {
    override val root = BorderPane()
    var sourceDir: TextField by singleAssign()
    var targetDir: TextField by singleAssign()
    var dictionaryGlobal: TextField by singleAssign()
    var dictionary: TextField by singleAssign()
    var languageFrom: TextField by singleAssign()
    var languageTo: TextField by singleAssign()
    var removeByColor: CheckBox by singleAssign()
    var removeUnusedFromGlobal: CheckBox by singleAssign()
    var colorToRemove: ColorPicker by singleAssign()


    val translateController: TranslateController by inject()

    var status: StringProperty = SimpleStringProperty()
    var statusUpdater: (String) -> Unit = { Platform.runLater { status.value = it } }

    init {
        title = "Translator"

        with(root) {
            setPrefSize(800.0, 200.0)
            top {
                vbox {
                    addClass(Styles.spaces)
                    hbox {
                        label("Source") {
                            hboxConstraints { margin = Insets(2.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        sourceDir = textfield() {
                            hboxConstraints {
                                margin = Insets(2.0)
                                hGrow = Priority.ALWAYS
                            }
                        }
                        button("...") {
                            hboxConstraints { margin = Insets(2.0) }
                            setOnAction {
                                val fileChooser = DirectoryChooser()
                                val selectedDirectory = fileChooser.showDialog(primaryStage)

                                if (selectedDirectory != null) {
                                    sourceDir.text = selectedDirectory.absolutePath
                                }
                            }
                        }
                    }
                    hbox {
                        label("Target") {
                            hboxConstraints { margin = Insets(2.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        targetDir = textfield() {
                            hboxConstraints {
                                margin = Insets(2.0)
                                hGrow = Priority.ALWAYS
                            }
                        }
                        button("...") {
                            hboxConstraints { margin = Insets(2.0) }
                            setOnAction {
                                val fileChooser = DirectoryChooser()
                                val selectedDirectory = fileChooser.showDialog(primaryStage)

                                if (selectedDirectory != null) {
                                    targetDir.text = selectedDirectory.absolutePath
                                }
                            }
                        }
                    }
                    hbox {
                        label("Dictionary Global") {
                            hboxConstraints { margin = Insets(2.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        dictionaryGlobal = textfield {
                            hboxConstraints {
                                margin = Insets(2.0)
                                hGrow = Priority.ALWAYS
                            }
                        }
                        button("...") {
                            hboxConstraints { margin = Insets(2.0) }
                            setOnAction {
                                val fileChooser = FileChooser()
                                val selectedDirectory = fileChooser.showOpenDialog(primaryStage)

                                if (selectedDirectory != null) {
                                    dictionaryGlobal.text = selectedDirectory.absolutePath
                                }
                            }
                        }
                    }
                    hbox {
                        label("Dictionary") {
                            hboxConstraints { margin = Insets(2.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        dictionary = textfield {
                            hboxConstraints {
                                margin = Insets(2.0)
                                hGrow = Priority.ALWAYS
                            }
                        }
                        button("...") {
                            hboxConstraints { margin = Insets(2.0) }
                            setOnAction {
                                val fileChooser = FileChooser()
                                val selectedDirectory = fileChooser.showOpenDialog(primaryStage)

                                if (selectedDirectory != null) {
                                    dictionary.text = selectedDirectory.absolutePath
                                }
                            }
                        }
                    }
                    hbox {
                        label("From") {
                            hboxConstraints { margin = Insets(2.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        languageFrom = textfield {
                            maxWidth = 50.0
                            hboxConstraints {
                                margin = Insets(2.0)
                                //hGrow = Priority.NEVER
                            }
                        }
                        label("To") {
                            hboxConstraints { margin = Insets(2.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        languageTo = textfield {
                            maxWidth = 50.0
                            hboxConstraints {
                                margin = Insets(2.0)
                                //hGrow = Priority.NEVER
                            }
                        }
                        label("Remove text by color") {
                            hboxConstraints { margin = Insets(2.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        removeByColor = checkbox {
                            hboxConstraints {
                                margin = Insets(2.0)
                                hGrow = Priority.NEVER
                            }
                        }
                        colorToRemove = colorpicker(Color.RED) {}
                        label("Remove unused from global") {
                            hboxConstraints { margin = Insets(2.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        removeUnusedFromGlobal = checkbox {
                            hboxConstraints {
                                margin = Insets(2.0)
                                hGrow = Priority.NEVER
                            }
                        }
                    }

                    hbox {
                        button("Translate") {
                            hboxConstraints { margin = Insets(2.0) }
                            setOnAction {
                                runAsync {
                                    translateController.translate()
                                } ui {
                                    statusUpdater("Done")
                                }
                            }
                        }

                        button("Clear") {
                            hboxConstraints { margin = Insets(2.0) }
                            setOnAction {
                                translateController.clearSettings()
                            }
                        }
                        button("Exit") {
                            hboxConstraints { margin = Insets(2.0) }
                            setOnAction {
                                translateController.exit()
                            }
                        }
                    }
                }
            }
            bottom {
                vbox {
                    separator()
                    label(status) {
                        hboxConstraints {
                            padding = Insets(2.0)
                            margin = Insets(2.0)
                            hGrow = Priority.ALWAYS
                        }
                    }
                }
            }
        }
    }
}