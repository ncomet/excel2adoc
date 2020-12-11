package fr.ncomet

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import picocli.CommandLine
import picocli.CommandLine.*
import picocli.CommandLine.Model.*
import java.io.File
import java.io.FileInputStream
import java.util.concurrent.Callable
import kotlin.system.exitProcess

val allowedFileExtensions = listOf("xlsx", "xls", "csv")

@Command(
    name = "excel2adoc",
    mixinStandardHelpOptions = true,
    version = ["excel2adoc v1.0"],
    description = ["Converts an .xlsx, .xls or .csv file into its asciidoc representation"]
)
class Excel2Asciidoc : Callable<Int> {

    @Spec lateinit var spec : CommandSpec

    @Parameters(index = "0", arity = "1", description=["An .xlsx, .xls or .csv file to convert to adoc"])
    lateinit var inputFile: File

    @Option(names = ["-o", "--output"], paramLabel = "example.adoc", description = ["the output file name, if not provided it will write to stdout"])
    var outputFileName: String? = null

    override fun call(): Int {
        when {
            !inputFile.exists() -> throw ParameterException(spec.commandLine(), "${inputFile.name} does not exist")
            inputFile.extension !in allowedFileExtensions -> throw ParameterException(spec.commandLine(), "${inputFile.name} needs to be of type ${allowedFileExtensions.joinToString()}")
            else -> {
                FileInputStream(inputFile).use {
                    val workBook = try {
                        XSSFWorkbook(it)
                    } catch (e: OLE2NotOfficeXmlFileException) {
                        HSSFWorkbook(it)
                    }
                }
                return ExitCode.OK
            }
        }
    }
}

internal fun command(args: Array<String>) = CommandLine(Excel2Asciidoc()).execute(*args)

fun main(args: Array<String>): Unit = exitProcess(command(args))