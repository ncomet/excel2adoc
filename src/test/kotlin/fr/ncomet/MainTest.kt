package fr.ncomet

import com.ginsberg.junit.exit.ExpectSystemExitWithStatus
import io.quarkus.test.junit.QuarkusTest
import org.junit.jupiter.api.Disabled
import org.junit.jupiter.api.Test

@QuarkusTest
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
    @Disabled
    fun `should load an xls file`() {
        command(arrayOf("src/test/resources/file.xls"))
    }

}