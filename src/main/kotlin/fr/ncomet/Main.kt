package fr.ncomet

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType.*
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import picocli.CommandLine
import picocli.CommandLine.*
import picocli.CommandLine.Help.Visibility.ALWAYS
import picocli.CommandLine.Model.CommandSpec
import java.io.File
import java.io.FileInputStream
import java.lang.System.lineSeparator
import java.util.concurrent.Callable
import kotlin.system.exitProcess

val allowedFileExtensions = listOf("xlsx", "xls", "csv")

fun main(args: Array<String>): Unit = exitProcess(command(args))

@Command(
    name = "excel2adoc",
    mixinStandardHelpOptions = true,
    version = ["excel2adoc v1.0"],
    description = ["Converts an .xlsx, .xls or .csv file into its asciidoc representation"]
)
class Excel2Asciidoc : Callable<Int> {

    @Spec
    lateinit var out: CommandSpec

    @Parameters(
        index = "0",
        description = ["An .xlsx, .xls or .csv file to convert to adoc"],
        arity = "1"
    )
    lateinit var inputFile: File

    @Option(
        names = ["-n", "--no-headers"],
        description = ["disables interpretation of first row as header"],
        defaultValue = "false",
        showDefaultValue = ALWAYS
    )
    var noHeaders: Boolean = false

    @Option(
        names = ["-s", "--sheet"],
        paramLabel = "2",
        description = ["sheet number, starting at 1. if not provided, will try to print all sheets"],
        arity = "1"
    )
    val sheetNumber: Int? = null

    override fun call(): Int {
        when {
            !inputFile.exists() -> throw ParameterException(out.commandLine(), "${inputFile.name} does not exist")
            inputFile.extension !in allowedFileExtensions -> throw ParameterException(
                out.commandLine(),
                "${inputFile.name} needs to be of type ${allowedFileExtensions.joinToString()}"
            )
            else -> {
                FileInputStream(inputFile).use {
                    val workBook = try {
                        XSSFWorkbook(it)
                    } catch (e: OLE2NotOfficeXmlFileException) {
                        HSSFWorkbook(it)
                    }
                    if (sheetNumber == null) {
                        workBook.forEach { sheet -> sheet.print(noHeaders) }
                    } else {
                        workBook.getSheetAt(sheetNumber - 1).print(noHeaders)
                    }
                }
                return ExitCode.OK
            }
        }
    }

    private fun Sheet.print(noHeaders: Boolean) {
        val maxOfFirstPhysicalCell = map { row -> row.firstCellNum.toInt() }.max() ?: 0
        val minOfFirstNonEmptyCell = map { row -> row.takeWhile { cell -> cell.isEmpty }.size }.min() ?: 0
        val emptyCellsToShift = maxOf(
            maxOfFirstPhysicalCell,
            minOfFirstNonEmptyCell
        )

        val rows = toList().takeLastWhile { row -> row.isNotEmpty }

        if (noHeaders) {
            rows.printColumnDescriptor(emptyCellsToShift)
            out.tableSeparator()
            rows.printContent(emptyCellsToShift)
        } else {
            out.tableSeparator()
            rows.printHeader(emptyCellsToShift)
            rows.printContentAfterHeader(emptyCellsToShift)
        }
        out.tableSeparator()
        out.newLine()
    }

    private fun List<Row>.printHeader(emptyCellsToShift: Int) =
        out.println(
            first().cellIterator()
                .asSequence()
                .filterIndexed { index, cell -> cell.isNotEmpty || index >= emptyCellsToShift }
                .map(renderCell)
                .joinToString(prefix = "|", separator = " |")
        )

    private fun List<Row>.printContentAfterHeader(emptyCellsToShift: Int) =
        out.println(
            drop(1).joinToString(separator = lineSeparator() + lineSeparator()) {
                it.cellIterator()
                    .asSequence()
                    .filterIndexed { index, cell -> cell.isNotEmpty || index >= emptyCellsToShift }
                    .map(renderCell)
                    .joinToString(prefix = "|", separator = "${lineSeparator()}|")
            }
        )

    private fun List<Row>.printContent(emptyCellsToShift: Int) =
        out.println(
            joinToString(separator = lineSeparator() + lineSeparator()) {
                it.cellIterator()
                    .asSequence()
                    .filterIndexed { index, cell -> cell.isNotEmpty || index >= emptyCellsToShift }
                    .map(renderCell)
                    .joinToString(prefix = "|", separator = "${lineSeparator()}|")
            }
        )

    private fun List<Row>.printColumnDescriptor(emptyCellsToShift: Int) =
        out.println(
            """[cols="${
                first().filterIndexed { index, cell -> cell.isNotEmpty || index >= emptyCellsToShift }.count()
            }*"]"""
        )
}


val renderCell: (Cell) -> String = { cell ->
    when (cell.cellType) {
        _NONE -> ""
        NUMERIC -> cell.numericCellValue.toString()
        STRING -> cell.stringCellValue
        FORMULA -> cell.cellFormula
        BLANK -> ""
        BOOLEAN -> cell.booleanCellValue.toString()
        ERROR -> cell.errorCellValue.toString()
        else -> ""
    }
}

internal fun command(args: Array<String>) = CommandLine(Excel2Asciidoc()).execute(*args)

private fun CommandSpec.tableSeparator() = commandLine().out.println("|===")
private fun CommandSpec.println(s: String) = commandLine().out.println(s)
private fun CommandSpec.newLine() = commandLine().out.println(lineSeparator())
private val Cell.isEmpty get() = cellType == BLANK || renderCell(this).trim() == ""
private val Cell.isNotEmpty get() = !isEmpty
private val Row.isNotEmpty get() = cellIterator().asSequence().any { it.isNotEmpty }