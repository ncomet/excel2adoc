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
            """|===
            :|First Name |Last Name |Gender |Country |Age |Date |Id |Number
            :|Dulce
            :|Abril
            :|Female
            :|United States
            :|32.0
            :|15/10/2017
            :|1562.0
            :|32.3
            :
            :|Mara
            :|Hashimoto
            :|Female
            :|Great Britain
            :|25.0
            :|16/08/2016
            :|1582.0
            :|23.0
            :
            :|Philip
            :|Gent
            :|Male
            :|France
            :|36.0
            :|21/05/2015
            :|2587.0
            :|15.5
            :
            :|Kathleen
            :|Hanner
            :|Female
            :|United States
            :|25.0
            :|15/10/2017
            :|3549.0
            :|23.5445
            :
            :|Nereida
            :|Magwood
            :|Female
            :|United States
            :|58.0
            :|16/08/2016
            :|2468.0
            :|11.0
            :
            :|Gaston
            :|Brumm
            :|Male
            :|United States
            :|24.0
            :|21/05/2015
            :|2554.0
            :|0.4455
            :
            :|Etta
            :|Hurn
            :|Female
            :|Great Britain
            :|56.0
            :|15/10/2017
            :|3598.0
            :|23.0
            :
            :|Earlean
            :|Melgar
            :|Female
            :|United States
            :|27.0
            :|16/08/2016
            :|2456.0
            :|14.0
            :
            :|Vincenza
            :|Weiland
            :|Female
            :|United States
            :|40.0
            :|21/05/2015
            :|6548.0
            :|6.0
            :|===
""".trimMargin(":")
        )
    }

}