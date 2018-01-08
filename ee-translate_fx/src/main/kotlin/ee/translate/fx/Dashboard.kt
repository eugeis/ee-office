package ee.translate.fx

import javafx.animation.KeyFrame
import javafx.animation.Timeline
import javafx.application.Platform
import javafx.beans.property.SimpleStringProperty
import javafx.beans.property.StringProperty
import javafx.collections.FXCollections
import javafx.event.EventHandler
import javafx.geometry.Insets
import javafx.geometry.Pos
import javafx.scene.control.*
import javafx.scene.layout.BorderPane
import javafx.scene.layout.Priority
import javafx.stage.DirectoryChooser
import javafx.stage.FileChooser
import javafx.util.Duration
import tornadofx.*
import java.io.File
import java.nio.file.Paths

class Dashboard : View() {
    override val root = BorderPane()
    val translateController: TranslateController by inject()
    var selectFrom: TextField by singleAssign()
    var interactive: CheckBox by singleAssign()
    var languageFrom: TextField by singleAssign()
    var languageTo: TextField by singleAssign()
    var dictionary: TextField by singleAssign()
    var filesTable: TableView<File>? = null
    var filesSelectionModel: TableView.TableViewSelectionModel<File>? = null
    var files = FXCollections.observableArrayList<File>()
    var translationTable: TableView<Map.Entry<String, String>>? = null
    var translationsSelectionModel: TableView.TableViewSelectionModel<Map.Entry<String, String>>? = null
    var translations = FXCollections.observableArrayList<Map.Entry<String, String>>(translateController.translations.entries)

    var translationsM = FXCollections.observableHashMap<String, String>()


    var status: StringProperty = SimpleStringProperty()
    var statusUpdater: (String) -> Unit = {
        Platform.runLater { status.value = it }
    }

    init {
        title = "Translator"

        with(root) {
            setPrefSize(800.0, 400.0)
            top {
                vbox {
                    addClass(Styles.spaces)
                    borderpane {
                        center {
                            hbox {
                                label("Select files") {
                                    hboxConstraints { margin = Insets(5.0) }
                                    alignment = Pos.CENTER_LEFT
                                    contentDisplay = ContentDisplay.LEFT
                                }
                                selectFrom = textfield() {
                                    hboxConstraints {
                                        margin = Insets(5.0)
                                        hGrow = Priority.ALWAYS
                                    }
                                }
                                button("...") {
                                    hboxConstraints { margin = Insets(5.0) }
                                    setOnAction {
                                        val directoryChooser = DirectoryChooser()
                                        val selectedDirectory = directoryChooser.showDialog(primaryStage)

                                        if (selectedDirectory != null) {
                                            selectFrom.text = selectedDirectory.absolutePath
                                            val pptxs = selectedDirectory.listFiles { file -> file.name.endsWith(".pptx") }
                                            files.addAll(pptxs)
                                        }
                                    }
                                }
                            }
                        }
                        right {
                            hbox {
                                button("Exit") {
                                    hboxConstraints { margin = Insets(5.0) }
                                    setOnAction {
                                        translateController.exit()
                                    }
                                }
                            }
                        }
                    }
                    hbox {
                        label("From") {
                            hboxConstraints { margin = Insets(5.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        languageFrom = textfield {
                            hboxConstraints {
                                margin = Insets(5.0)
                                hGrow = Priority.NEVER
                            }
                            setOnAction {
                                translateController.changeFrom(text)
                            }
                        }
                        label("To") {
                            hboxConstraints { margin = Insets(5.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        languageTo = textfield {
                            hboxConstraints {
                                margin = Insets(5.0)
                                hGrow = Priority.NEVER
                            }
                            setOnAction {
                                translateController.changeTo(text)
                            }
                        }
                        label("Dictionary") {
                            hboxConstraints { margin = Insets(5.0) }
                            alignment = Pos.CENTER_LEFT
                            contentDisplay = ContentDisplay.LEFT
                        }
                        dictionary = textfield {
                            hboxConstraints {
                                margin = Insets(5.0)
                                hGrow = Priority.ALWAYS
                            }
                        }
                        button("...") {
                            hboxConstraints { margin = Insets(5.0) }
                            setOnAction {
                                val fileChooser = FileChooser()
                                val selectedDirectory = fileChooser.showOpenDialog(primaryStage)

                                if (selectedDirectory != null) {
                                    dictionary.text = selectedDirectory.absolutePath
                                    translateController.changeDictionary(Paths.get(dictionary.text))
                                }
                            }
                        }
                    }
                    hbox {
                        vboxConstraints { marginTop = 5.0 }
                        hbox {
                            label("Interactive") {
                                hboxConstraints { margin = Insets(5.0) }
                            }
                            interactive = checkbox() {
                                hboxConstraints { margin = Insets(5.0) }
                            }
                        }
                        button("Clear") {
                            hboxConstraints { margin = Insets(5.0) }
                            setOnAction {
                                translateController.clearSettings()
                            }
                        }
                        button("Translate") {
                            hboxConstraints { margin = Insets(5.0) }
                            setOnAction {
                                translateController.translate(filesSelectionModel?.selectedItems!!,
                                        interactive.isSelected, statusUpdater)
                            }
                        }
                    }
                }
            }
            center {
                hbox {
                    filesTable = tableview(files) {
                        filesSelectionModel = selectionModel
                        selectionModel.selectionMode = SelectionMode.MULTIPLE
                        column("Files", File::getName).remainingWidth()
                        columnResizePolicy = SmartResize.POLICY
                    }
                    translationTable = tableview(translations) {
                        translationsSelectionModel = selectionModel
                        selectionModel.selectionMode = SelectionMode.MULTIPLE
                        column("from", Map.Entry<String, String>::key).remainingWidth()
                        column("to", Map.Entry<String, String>::value).remainingWidth()
                        columnResizePolicy = SmartResize.POLICY
                    }
                }
            }
            bottom {
                vbox {
                    separator()
                    label(status) {
                        hboxConstraints {
                            padding = Insets(2.0)
                            margin = Insets(5.0)
                            hGrow = Priority.ALWAYS
                        }
                    }
                }
            }
        }

    }

    fun shakeStage() {
        var x = 0
        var y = 0
        val cycleCount = 10
        val move = 10
        val keyframeDuration = Duration.seconds(0.04)

        val stage = FX.primaryStage

        val timelineX = Timeline(KeyFrame(keyframeDuration, EventHandler {
            if (x == 0) {
                stage.x = stage.x + move
                x = 1
            } else {
                stage.x = stage.x - move
                x = 0
            }
        }))

        timelineX.cycleCount = cycleCount
        timelineX.isAutoReverse = false

        val timelineY = Timeline(KeyFrame(keyframeDuration, EventHandler {
            if (y == 0) {
                stage.y = stage.y + move
                y = 1;
            } else {
                stage.y = stage.y - move
                y = 0;
            }
        }))

        timelineY.cycleCount = cycleCount;
        timelineY.isAutoReverse = false;

        timelineX.play()
        timelineY.play();
    }
}