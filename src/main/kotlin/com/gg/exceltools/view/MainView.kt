package com.gg.exceltools.view

import com.alibaba.excel.EasyExcel
import com.alibaba.excel.context.AnalysisContext
import com.alibaba.excel.event.AnalysisEventListener
import com.gg.exceltools.data.Version
import javafx.geometry.Pos
import javafx.scene.control.TextField
import kotlinx.coroutines.GlobalScope
import kotlinx.coroutines.launch
import tornadofx.*
import java.io.File
import java.lang.Exception
import java.text.SimpleDateFormat
import java.util.*
import javax.swing.JFileChooser


class MainView : View("${Version.name} (${Version.code})") {

    private var firstExcelPathStr: String = ""
    private var secondExcelPathStr: String = ""
    private var firstFixedDataList: List<Int> = listOf()
    private var secondFixedDataList: List<Int> = listOf()
    private var firstCompareDataList: List<Int> = listOf()
    private var secondCompareDataList: List<Int> = listOf()
    private var firstNameStr: String = ""
    private var secondNameStr: String = ""
    private var firstDateInt: Int = 0
    private var secondDateInt: Int = 0
    private var firstStartInt: Int = 0
    private var secondStartInt: Int = 0

    private val file = File("")

    private var firstExcelPathEdit: TextField = TextField()
    private var secondExcelPathEdit: TextField = TextField()
    private var firstFixedDataEdit: TextField = TextField()
    private var secondFixedDataEdit: TextField = TextField()
    private var firstDateEdit: TextField = TextField()
    private var secondDateEdit: TextField = TextField()
    private var firstCompareDataEdit: TextField = TextField()
    private var secondCompareDataEdit: TextField = TextField()
    private var firstNameEdit: TextField = TextField()
    private var secondNameEdit: TextField = TextField()
    private var firstStartEdit: TextField = TextField()
    private var secondStartEdit: TextField = TextField()


    override val root = vbox(alignment = Pos.CENTER) {

        paddingAll = 20.0

        form {
            fieldset("选择Excel文件") {
                field("第一个Excel文件") {

                    firstExcelPathEdit = textfield()

                    button("选择文件").action {
                        val fileChooser = JFileChooser(file.absolutePath);

                        fileChooser.fileSelectionMode = JFileChooser.FILES_ONLY;

                        val returnVal = fileChooser.showOpenDialog(fileChooser);

                        if (returnVal == JFileChooser.APPROVE_OPTION) {
                            val filePath = fileChooser.selectedFile.absolutePath;//这个就是你选择的文件的
                            firstExcelPathEdit.text = filePath
                        }
                    }
                }

                field("第二个Excel文件") {

                    secondExcelPathEdit = textfield()
                    button("选择文件").action {
                        val fileChooser = JFileChooser(file.absolutePath);

                        fileChooser.fileSelectionMode = JFileChooser.FILES_ONLY;

                        val returnVal = fileChooser.showOpenDialog(fileChooser);

                        if (returnVal == JFileChooser.APPROVE_OPTION) {
                            val filePath = fileChooser.selectedFile.absolutePath;//这个就是你选择的文件的
                            secondExcelPathEdit.text = filePath

                        }
                    }
                }
            }

            fieldset("固定参数位置(多个参数一一对应,用 \",\" 隔开)") {
                field("第一个Excel文件的参数位置") {
                    firstFixedDataEdit = textfield()
                }
                field("第二个Excel文件的参数位置") {
                    secondFixedDataEdit = textfield()
                }
            }
            fieldset("比较参数位置(多个参数一一对应,用 \",\" 隔开)") {
                field("第一个Excel文件的参数位置") {
                    firstCompareDataEdit = textfield()
                }
                field("第二个Excel文件的参数位置") {
                    secondCompareDataEdit = textfield()
                }
            }
            fieldset("时间参数位置") {
                field("第一个Excel文件的时间参数位置") {
                    firstDateEdit = textfield()
                }
                field("第二个Excel文件的时间参数位置") {
                    secondDateEdit = textfield()
                }
            }

            fieldset("选择两个Excel中需要对比的表") {
                field("第一个Excel中的表名称") {
                    firstNameEdit = textfield()
                }
                field("第二个Excel中的表名称") {
                    secondNameEdit = textfield()
                }
            }

            fieldset("从两个文件的第几行开始比较") {
                field("第一个Excel文件的开始行数") {
                    firstStartEdit = textfield()
                }
                field("第二个Excel文件的开始行数") {
                    secondStartEdit = textfield()
                }

            }
        }

        button("开始比较并输出文件") {
            paddingAll = 10.0
        }.action {
            try {
                firstStartInt = firstStartEdit.text.toInt() - 1
                secondStartInt = secondStartEdit.text.toInt() - 1
                firstExcelPathStr = firstExcelPathEdit.text
                secondExcelPathStr = secondExcelPathEdit.text
                firstNameStr = firstNameEdit.text
                secondNameStr = secondNameEdit.text

                firstDateInt = firstDateEdit.text.toInt() - 1
                secondDateInt = secondDateEdit.text.toInt() - 1

                firstFixedDataList = firstFixedDataEdit.text.split(",").map { it.toInt() - 1 }
                secondFixedDataList = secondFixedDataEdit.text.split(",").map { it.toInt() - 1 }

                firstCompareDataList = firstCompareDataEdit.text.split(",").map { it.toInt() - 1 }
                secondCompareDataList = secondCompareDataEdit.text.split(",").map { it.toInt() - 1 }
                if (firstFixedDataList.size != secondFixedDataList.size || firstCompareDataList.size != secondCompareDataList.size)
                    throw Exception()

            } catch (e: Exception) {
                ToastUtil.toast(msg = "输入参数有误，请重新选择或填入")
                return@action
            }

            if (firstExcelPathStr.isEmpty() || !firstExcelPathStr.isExcelEnd() ||
                    secondExcelPathStr.isEmpty() || !secondExcelPathStr.isExcelEnd() ||
                    firstFixedDataList.isEmpty() || secondFixedDataList.isEmpty() ||
                    firstNameStr.isEmpty() || secondNameStr.isEmpty() ||
                    firstStartInt < 0 || secondStartInt < 0 ||
                    !firstFixedDataList.none { it < 0 } || !secondFixedDataList.none { it < 0 } ||
                    !firstCompareDataList.none { it < 0 } || !secondCompareDataList.none { it < 0 } ||
                    firstDateInt < 0 || secondDateInt < 0
            ) {

                ToastUtil.toast(msg = "输入参数有误，请重新选择或填入")
                return@action
            } else {
                parseData()
            }
        }
    }

    private fun parseData() {

        GlobalScope.launch {
            //读取数据
            val firstList = EasyExcel.read(firstExcelPathStr).sheet(firstNameStr).headRowNumber(firstStartInt).doReadSync<Map<Int, String>>()
            val secondList = EasyExcel.read(secondExcelPathStr).sheet(secondNameStr).headRowNumber(secondStartInt).doReadSync<Map<Int, String>>()

            val dataList = mutableListOf<List<String>>()

            var tempList: MutableList<String> = mutableListOf()
            //填入标题头
            firstFixedDataList.forEach {
                tempList.add(firstList[0].values.toList()[it])
            }
            println(firstList[0])
            tempList.add(firstList[0].values.toList()[firstDateInt])
            firstCompareDataList.forEach {
                tempList.add(firstList[0].values.toList()[it])
            }
            dataList.add(tempList)

            //选出需要比较的对象不一致的数据 存储到 dataList 中
            firstList.filter {
                it.isNotEmpty()
            }.forEachIndexed { _, firstMap ->
                secondList.forEachIndexed { _, secondMap ->

                    var firstDateList = firstMap[firstDateInt]?.split("/")?.filter { it.isNotEmpty() }
                    var secondDateList = secondMap[secondDateInt]?.split("/")?.filter { it.isNotEmpty() }

                    firstDateList = parseDate(firstDateList)
                    secondDateList = parseDate(secondDateList)

                    if (firstFixedDataList.mapIndexed { index, value ->
                                return@mapIndexed firstMap[value] == secondMap[secondFixedDataList[index]]
                            }.all {
                                it
                            } && firstDateList?.filter { it.length <= 2 }.toString() == secondDateList?.filter { it.length <= 2 }.toString()) {
                        if (!firstCompareDataList.mapIndexed { index, value ->
                                    try {
                                        return@mapIndexed firstMap[value]?.toFloat() == secondMap[secondCompareDataList[index]]?.toFloat()
                                    } catch (e: Exception) {

                                        return@mapIndexed firstMap[value] == secondMap[secondCompareDataList[index]]
                                    }
                                }.all {
                                    it
                                }) {
                            //重新组装数据
                            tempList = mutableListOf()
                            firstFixedDataList.forEach {
                                tempList.add(firstMap.values.toList()[it])
                            }
                            tempList.add(firstMap.values.toList()[firstDateInt])

                            firstCompareDataList.forEach {
                                tempList.add(firstMap.values.toList()[it])
                            }
                            dataList.add(tempList)

                            tempList = mutableListOf()
                            secondFixedDataList.forEach {
                                tempList.add(secondMap.values.toList()[it])
                            }
                            tempList.add(secondMap.values.toList()[secondDateInt])
                            secondCompareDataList.forEach {
                                tempList.add(secondMap.values.toList()[it])
                            }
                            dataList.add(tempList)
                            dataList.add(mutableListOf())
                        }
                    }
                }

            }

            //写入数据到指定位置
            val fileName = File("").absolutePath + File.separator + "比较结果${SimpleDateFormat("yyyy-MM-dd_HH-mm-ss").format(Date())}.xlsx"
            EasyExcel.write(fileName).sheet("比较结果").doWrite(dataList)
            println("导出完成")

        }

        ToastUtil.toast(msg = "导出完成")
    }

    fun parseDate(list: List<String>?): List<String>? {
        if (list?.size == 1) {
            return list[0].split(" ")[0].split("-").filter { it.length <= 2 }
        } else {
            return list
        }
    }

    private fun String.isExcelEnd(): Boolean {
        return this.endsWith("xls") || this.endsWith("xlsx")
    }
}