package fr.ncomet

import com.ginsberg.junit.exit.ExpectSystemExitWithStatus
import org.assertj.core.api.Assertions.assertThat
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
    //@Disabled(value = "to check: NotOLE2FileException: Invalid header signature; read 0x006D007500530005, expected 0xE11AB1A1E011CFD0 - Your file appears not to be a valid OLE2 document, maybe a maven resource issue ?")
    @ExpectSystemExitWithStatus(2)
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

    @Test
    @StdIo
    fun `should produce expected output on a shifted xlsx file`(out: StdOut) {
        command(arrayOf("src/test/resources/file-shifted.xlsx"))

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

    @Test
    @StdIo
    fun `should produce expected output on an xlsx file without header`(out: StdOut) {
        command(arrayOf("src/test/resources/file.xlsx", "-n"))

        assertThat(out.capturedLines().joinToString(System.lineSeparator())).isEqualTo(
            """[cols="8*"]
            :|===
            :|First Name
            :|Last Name
            :|Gender
            :|Country
            :|Age
            :|Date
            :|Id
            :|Number
            :
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

    @Test
    @StdIo
    fun `should produce expected output on a shifted xlsx file without header`(out: StdOut) {
        command(arrayOf("src/test/resources/file-shifted.xlsx", "-n"))

        assertThat(out.capturedLines().joinToString(System.lineSeparator())).isEqualTo(
            """[cols="8*"]
            :|===
            :|First Name
            :|Last Name
            :|Gender
            :|Country
            :|Age
            :|Date
            :|Id
            :|Number
            :
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

    @Test
    @StdIo
    fun `should produce expected output on an xlsx file with 3 sheets`(out: StdOut) {
        command(arrayOf("src/test/resources/file-3-sheets.xlsx"))

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
            :|===
            :
            :
            :|===
            :|Animal |Weight |Temperature
            :|Duck
            :|1.3
            :|180.0
            :
            :|Cow
            :|1250.0
            :|200.0
            :
            :|Lamb
            :|30.0
            :|175.0
            :|===
            :
            :
            :|===
            :|A |B |C |D |E
            :|1.0
            :|2.0
            :|3.0
            :|4.0
            :|5.0
            :
            :|6.0
            :|7.0
            :|8.0
            :|9.0
            :|10.0
            :
            :|11.0
            :|12.0
            :|13.0
            :|14.0
            :|15.0
            :
            :|16.0
            :|17.0
            :|18.0
            :|19.0
            :|20.0
            :
            :|21.0
            :|22.0
            :|23.0
            :|24.0
            :|25.0
            :|===
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should produce expected output on a shifted xlsx file with 3 sheets`(out: StdOut) {
        command(arrayOf("src/test/resources/file-3-sheets-shifted.xlsx"))

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
            :|===
            :
            :
            :|===
            :|Animal |Weight |Temperature
            :|Duck
            :|1.3
            :|180.0
            :
            :|Cow
            :|1250.0
            :|200.0
            :
            :|Lamb
            :|30.0
            :|175.0
            :|===
            :
            :
            :|===
            :|A |B |C |D |E
            :|1.0
            :|2.0
            :|3.0
            :|4.0
            :|5.0
            :
            :|6.0
            :|7.0
            :|8.0
            :|9.0
            :|10.0
            :
            :|11.0
            :|12.0
            :|13.0
            :|14.0
            :|15.0
            :
            :|16.0
            :|17.0
            :|18.0
            :|19.0
            :|20.0
            :
            :|21.0
            :|22.0
            :|23.0
            :|24.0
            :|25.0
            :|===
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should produce expected output on an xlsx file with 3 sheets with the -s option`(out: StdOut) {
        command(arrayOf("src/test/resources/file-3-sheets.xlsx", "-s=2"))

        assertThat(out.capturedLines().joinToString(System.lineSeparator())).isEqualTo(
            """|===
            :|Animal |Weight |Temperature
            :|Duck
            :|1.3
            :|180.0
            :
            :|Cow
            :|1250.0
            :|200.0
            :
            :|Lamb
            :|30.0
            :|175.0
            :|===
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should produce expected output on a shifted xlsx file with 3 sheets with the -s option`(out: StdOut) {
        command(arrayOf("src/test/resources/file-3-sheets-shifted.xlsx", "-s=2"))

        assertThat(out.capturedLines().joinToString(System.lineSeparator())).isEqualTo(
            """|===
            :|Animal |Weight |Temperature
            :|Duck
            :|1.3
            :|180.0
            :
            :|Cow
            :|1250.0
            :|200.0
            :
            :|Lamb
            :|30.0
            :|175.0
            :|===
""".trimMargin(":")
        )
    }
}