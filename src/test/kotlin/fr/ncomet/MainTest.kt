package fr.ncomet

import com.ginsberg.junit.exit.ExpectSystemExitWithStatus
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Disabled
import org.junit.jupiter.api.Test
import org.junitpioneer.jupiter.StdIo
import org.junitpioneer.jupiter.StdOut

class MainTest {

    @Test
    @ExpectSystemExitWithStatus(2)
    fun `should exit(2) if file does not exists`() {
        main(arrayOf("src/test/resources/unknown.xlsx"))
    }

    @Test
    @ExpectSystemExitWithStatus(2)
    fun `should exit(2) if file extension is not allowed`() {
        main(arrayOf("src/test/resources/file.test"))
    }

    @Test
    fun `should load an xlsx file`() {
        command(arrayOf("src/test/resources/file.xlsx"))
    }

    @Test
    @Disabled(value = "to check: NotOLE2FileException: Invalid header signature; read 0x006D007500530005, expected 0xE11AB1A1E011CFD0 - Your file appears not to be a valid OLE2 document")
    fun `should load an xls file`() {
        command(arrayOf("src/test/resources/file.xls"))
    }

    @Test
    @StdIo
    fun `should produce expected output on an xlsx file`(out: StdOut) {
        command(arrayOf("src/test/resources/file.xlsx"))

        assertThat(out.capturedLines().joinToString(System.lineSeparator())).isEqualTo(
            """
            :|===
            :|Header 1 |Header 2 |Header 3 |Header 4
            :
            :|Column 1, row 1
            :|Column 2, row 1
            :|Column 3, row 1
            :|Column 4, row 1
            :
            :|Column 1, row 2
            :|Column 2, row 2
            :|Column 3, row 2
            :|Column 4, row 2
            :
            :|Column 1, row 3
            :|Column 2, row 3
            :|Column 3, row 3
            :|Column 4, row 3
            :|===
:""".trimMargin(":")
        )
    }

}