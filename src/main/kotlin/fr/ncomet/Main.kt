package fr.ncomet

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
    version = ["excel2adoc 1.0"],
    description = ["converts an .xlsx, .xls or .csv file into its asciidoc representation"]
)
class Excel2Asciidoc : Callable<Int> {

    @Spec lateinit var spec : CommandSpec

    @Parameters(index = "0", arity = "1", description=["An .xlsx, .xls or .csv file to convert to adoc"])
    lateinit var inputFile: File

    @Option(names = ["-o", "--output"], paramLabel = "example.adoc", description = ["the output file name, id not provided it will default to stdout"])
    var outputFileName: String? = null

    override fun call(): Int {
        when {
            !inputFile.exists() -> throw ParameterException(spec.commandLine(), "${inputFile.name} does not exist")
            inputFile.extension !in allowedFileExtensions -> throw ParameterException(spec.commandLine(), "${inputFile.name} needs to be of type ${allowedFileExtensions.joinToString()}")
            else -> {
                println("file: $inputFile and name: $outputFileName")
                FileInputStream(inputFile).use {
                    val xssfWorkbook = XSSFWorkbook(it)
                }
                return ExitCode.OK
            }
        }
    }
}

fun main(args: Array<String>): Unit = exitProcess(CommandLine(Excel2Asciidoc()).execute(*args))