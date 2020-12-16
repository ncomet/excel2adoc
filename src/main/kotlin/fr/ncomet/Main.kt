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

val allowedFileExtensions = listOf("xlsx", "csv")

fun main(args: Array<String>): Unit = exitProcess(command(args))

@Command(
    name = "excel2adoc",
    mixinStandardHelpOptions = true,
    version = ["excel2adoc v1.0"],
    description = ["Prints .xlsx or .csv file(s) into their asciidoc representation on stdout"]
)
class Excel2Asciidoc : Callable<Int> {

    @Spec
    lateinit var out: CommandSpec

    @Parameters(
        index = "0",
        description = ["One or more .xlsx or .csv file(s) to print to stdout"],
        arity = "1..*"
    )
    lateinit var inputFiles: List<File>

    @Option(
        names = ["-n", "--no-headers"],
        description = ["Disables interpretation of first row as header"],
        defaultValue = "false",
        showDefaultValue = ALWAYS
    )
    var noHeaders: Boolean = false

    @Option(
        names = ["-t", "--no-titles"],
        description = ["Disables table title using file name and sheet name (or just file name for .csv)"],
        defaultValue = "false",
        showDefaultValue = ALWAYS
    )
    var noTitles: Boolean = false

    @Option(
        names = ["-s", "--sheet"],
        paramLabel = "2",
        description = ["Sheet number, starting at 1. if not provided, it will try to print all sheets for all files"],
        arity = "1"
    )
    val sheetNumber: Int? = null

    override fun call(): Int {
        inputFiles.forEach {
            if (!it.exists()) throw ParameterException(out.commandLine(), "${it.name} does not exist")
        }

        inputFiles.forEach {
            if (it.extension !in allowedFileExtensions) throw ParameterException(
                out.commandLine(),
                "${it.name} needs to be of type ${allowedFileExtensions.joinToString()}"
            )
        }

        inputFiles.forEach { file ->
            when (file.extension) {
                "xlsx" -> FileInputStream(file).use {
                    val workBook = try {
                        XSSFWorkbook(it)
                    } catch (e: OLE2NotOfficeXmlFileException) {
                        HSSFWorkbook(it)
                    }
                    if (sheetNumber == null) {
                        workBook.forEach { sheet -> sheet.print(noHeaders, file.nameWithoutExtension) }
                    } else {
                        workBook.getSheetAt(sheetNumber - 1).print(noHeaders, file.nameWithoutExtension)
                    }
                }
                else -> with(file) { print() }
            }
            out.newLine()
        }
        return ExitCode.OK
    }

    private fun File.print() {
        val fileLines = readLines()
        if (fileLines.isNotEmpty()) {
            out.println("[${if (noHeaders) "" else "%header,"}format=csv,separator=;]")
            if (!noTitles) out.println(".${nameWithoutExtension}")
            out.println("|===")
            out.println(fileLines.joinToString(separator = lineSeparator()))
            out.println("|===")
        }
    }

    private fun Sheet.print(noHeaders: Boolean, fileName: String) {
        val maxOfFirstPhysicalCell = map { row -> row.firstCellNum.toInt() }.max() ?: 0
        val minOfFirstNonEmptyCell = map { row -> row.takeWhile { cell -> cell.isEmpty }.size }.min() ?: 0
        val emptyCellsToShift = maxOf(maxOfFirstPhysicalCell, minOfFirstNonEmptyCell)

        val rows = toList().takeLastWhile { row -> row.isNotEmpty }

        if (rows.isNotEmpty()) {
            if (noHeaders) {
                rows.printColumnDescriptor(emptyCellsToShift)
                tableSeparator(fileName)
                rows.printContent(emptyCellsToShift)
            } else {
                tableSeparator(fileName)
                rows.printHeader(emptyCellsToShift)
                rows.printContentAfterHeader(emptyCellsToShift)
            }
            out.println("|===")
            out.newLine()
        }
    }

    private fun List<Row>.printColumnDescriptor(emptyCellsToShift: Int) =
        out.println(
            """[cols="${
                first().filterIndexed { index, cell -> cell.isNotEmpty || index >= emptyCellsToShift }.count()
            }*"]"""
        )

    private fun Sheet.tableSeparator(fileName: String) =
        out.println(if (noTitles) "|===" else ".$fileName $sheetName${lineSeparator()}|===")

    private fun List<Row>.printHeader(emptyCellsToShift: Int) =
        out.println(
            first().cellIterator()
                .asSequence()
                .filterIndexed { index, cell -> cell.isNotEmpty || index >= emptyCellsToShift }
                .map(renderCell)
                .joinToString(prefix = "|", separator = " |")
        )

    private fun List<Row>.printContentAfterHeader(emptyCellsToShift: Int) =
        out.println(drop(1).joinRows(emptyCellsToShift))

    private fun List<Row>.printContent(emptyCellsToShift: Int) =
        out.println(joinRows(emptyCellsToShift))

    private fun List<Row>.joinRows(emptyCellsToShift: Int) =
        joinToString(separator = lineSeparator() + lineSeparator()) {
            it.cellIterator()
                .asSequence()
                .filterIndexed { index, cell -> cell.isNotEmpty || index >= emptyCellsToShift }
                .map(renderCell)
                .joinToString(prefix = "|", separator = "${lineSeparator()}|")
        }

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
private fun CommandSpec.println(s: String) = commandLine().out.println(s)
private fun CommandSpec.newLine() = commandLine().out.println(lineSeparator())
private val Cell.isEmpty get() = cellType == BLANK || renderCell(this).trim() == ""
private val Cell.isNotEmpty get() = !isEmpty
private val Row.isNotEmpty get() = cellIterator().asSequence().any { it.isNotEmpty }