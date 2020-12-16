package fr.ncomet

import com.ginsberg.junit.exit.ExpectSystemExitWithStatus
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Disabled
import org.junit.jupiter.api.Test
import org.junitpioneer.jupiter.StdIo
import org.junitpioneer.jupiter.StdOut
import java.lang.System.lineSeparator

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
    @StdIo
    fun `should produce expected output on an xlsx file`(out: StdOut) {
        command(arrayOf("src/test/resources/file.xlsx"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """.file Sheet1
            :|===
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

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """.file-shifted Sheet1
            :|===
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
    fun `should produce expected output on an xlsx file without header (-n option)`(out: StdOut) {
        command(arrayOf("src/test/resources/file.xlsx", "-n"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """[cols="8*"]
            :.file Sheet1
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
    fun `should produce expected output on a shifted xlsx file without header (-n option)`(out: StdOut) {
        command(arrayOf("src/test/resources/file-shifted.xlsx", "-n"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """[cols="8*"]
            :.file-shifted Sheet1
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

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """.file-3-sheets Sheet1
            :|===
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
            :.file-3-sheets Sheet2
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
            :.file-3-sheets Sheet3
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

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """.file-3-sheets-shifted Sheet1
            :|===
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
            :.file-3-sheets-shifted Sheet2
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
            :.file-3-sheets-shifted Sheet3
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
    fun `should produce expected output on an xlsx file with 3 sheets with explicit sheet (-s option)`(out: StdOut) {
        command(arrayOf("src/test/resources/file-3-sheets.xlsx", "-s=2"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """.file-3-sheets Sheet2
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
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should produce expected output without titles (-t option)`(out: StdOut) {
        command(arrayOf("src/test/resources/file-3-sheets.xlsx", "-s=2", "-t"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
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
    fun `should produce expected output on a shifted xlsx file with 3 sheets with explicit sheet (-s option)`(out: StdOut) {
        command(arrayOf("src/test/resources/file-3-sheets-shifted.xlsx", "-s=2"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """.file-3-sheets-shifted Sheet2
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
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should print empty output for an empty xlsx file`(out: StdOut) {
        command(arrayOf("src/test/resources/empty-file.xlsx"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo("")
    }

    @Test
    @StdIo
    fun `should product output for every cell type`(out: StdOut) {
        command(arrayOf("src/test/resources/file-every-cell-types.xlsx"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """.file-every-cell-types Feuille1
            :|===
            :|String |1.0 |1.65 |SUM(B1,C1) |wtf() |1/0 |true
            :
            :|===
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should process a csv file`(out: StdOut) {
        command(arrayOf("src/test/resources/file.csv"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """[%header,format=csv,separator=;]
            :.file
            :|===
            :First Name;Last Name;Gender;Country;Age;Date;Id;Number
            :Dulce;Abril;Female;United States;32;15/10/2017;1562;32
            :Mara;Hashimoto;Female;Great Britain;25;16/08/2016;1582;23
            :Philip;Gent;Male;France;36;21/05/2015;2587;16
            :Kathleen;Hanner;Female;United States;25;15/10/2017;3549;24
            :Nereida;Magwood;Female;United States;58;16/08/2016;2468;11
            :Gaston;Brumm;Male;United States;24;21/05/2015;2554;0
            :Etta;Hurn;Female;Great Britain;56;15/10/2017;3598;23
            :Earlean;Melgar;Female;United States;27;16/08/2016;2456;14
            :Vincenza;Weiland;Female;United States;40;21/05/2015;6548;6
            :|===
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should process a csv file without header (-n option)`(out: StdOut) {
        command(arrayOf("src/test/resources/file.csv", "-n"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """[format=csv,separator=;]
            :.file
            :|===
            :First Name;Last Name;Gender;Country;Age;Date;Id;Number
            :Dulce;Abril;Female;United States;32;15/10/2017;1562;32
            :Mara;Hashimoto;Female;Great Britain;25;16/08/2016;1582;23
            :Philip;Gent;Male;France;36;21/05/2015;2587;16
            :Kathleen;Hanner;Female;United States;25;15/10/2017;3549;24
            :Nereida;Magwood;Female;United States;58;16/08/2016;2468;11
            :Gaston;Brumm;Male;United States;24;21/05/2015;2554;0
            :Etta;Hurn;Female;Great Britain;56;15/10/2017;3598;23
            :Earlean;Melgar;Female;United States;27;16/08/2016;2456;14
            :Vincenza;Weiland;Female;United States;40;21/05/2015;6548;6
            :|===
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should process a csv file without titles (-t option)`(out: StdOut) {
        command(arrayOf("src/test/resources/file.csv", "-t"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """[%header,format=csv,separator=;]
            :|===
            :First Name;Last Name;Gender;Country;Age;Date;Id;Number
            :Dulce;Abril;Female;United States;32;15/10/2017;1562;32
            :Mara;Hashimoto;Female;Great Britain;25;16/08/2016;1582;23
            :Philip;Gent;Male;France;36;21/05/2015;2587;16
            :Kathleen;Hanner;Female;United States;25;15/10/2017;3549;24
            :Nereida;Magwood;Female;United States;58;16/08/2016;2468;11
            :Gaston;Brumm;Male;United States;24;21/05/2015;2554;0
            :Etta;Hurn;Female;Great Britain;56;15/10/2017;3598;23
            :Earlean;Melgar;Female;United States;27;16/08/2016;2456;14
            :Vincenza;Weiland;Female;United States;40;21/05/2015;6548;6
            :|===
""".trimMargin(":")
        )
    }

    @Test
    @StdIo
    fun `should print empty output for an empty csv file`(out: StdOut) {
        command(arrayOf("src/test/resources/empty-file.csv"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo("")
    }

    @Test
    @StdIo
    fun `should print two files`(out: StdOut) {
        command(arrayOf("src/test/resources/file.csv", "src/test/resources/file-3-sheets.xlsx"))

        assertThat(out.capturedLines().joinToString(lineSeparator())).isEqualTo(
            """[%header,format=csv,separator=;]
            :.file
            :|===
            :First Name;Last Name;Gender;Country;Age;Date;Id;Number
            :Dulce;Abril;Female;United States;32;15/10/2017;1562;32
            :Mara;Hashimoto;Female;Great Britain;25;16/08/2016;1582;23
            :Philip;Gent;Male;France;36;21/05/2015;2587;16
            :Kathleen;Hanner;Female;United States;25;15/10/2017;3549;24
            :Nereida;Magwood;Female;United States;58;16/08/2016;2468;11
            :Gaston;Brumm;Male;United States;24;21/05/2015;2554;0
            :Etta;Hurn;Female;Great Britain;56;15/10/2017;3598;23
            :Earlean;Melgar;Female;United States;27;16/08/2016;2456;14
            :Vincenza;Weiland;Female;United States;40;21/05/2015;6548;6
            :|===
            :
            :
            :.file-3-sheets Sheet1
            :|===
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
            :.file-3-sheets Sheet2
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
            :.file-3-sheets Sheet3
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
}