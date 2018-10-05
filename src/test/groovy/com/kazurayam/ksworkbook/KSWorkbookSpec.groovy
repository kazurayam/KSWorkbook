package com.kazurayam.ksworkbook

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import java.time.LocalDateTime
import java.time.ZoneId

import org.junit.Test
import org.slf4j.Logger
import org.slf4j.LoggerFactory

import spock.lang.Specification

class KSWorkbookSpec extends Specification {

    static Logger logger_ = LoggerFactory.getLogger(KSWorkbookSpec.class)
    
    // fields
    private static Path workdir_
    
    // fixture methods
    
    def setupSpec() {
        workdir_ = Paths.get("./build/tmp/${Helpers.getClassShortName(KSWorkbookSpec.class)}")
        if (!Files.exists(workdir_)) {
            Files.createDirectories(workdir_)
        }
    }
    def setup() {}
    def cleanup() {}
    def cleanupSpec() {}
    
    
    def testSmoke() {
        when:
        Path excel = workdir_.resolve("testSmoke.xls")
        KSWorkbook wb = new KSWorkbook(excel)
        then:
        wb != null
        when:
        wb.writeCell("Sheet1", 0, 0, "Hello")
        wb.close()
        then:
        Files.exists(excel)
    }
    
    def testCreatingMultipleSheets() {
        when:
        Path excel = workdir_.resolve("testCreatingMultipleSheets.xls")
        KSWorkbook wb = new KSWorkbook(excel)
        wb.writeCell("Sheet1", 0, 0, "Good Morning")
        wb.writeCell("Sheet2", 0, 0, "Good Afternoon")
        wb.writeCell("Sheet3", 0, 0, "Good Evening")
        wb.close()
        then:
        Files.exists(excel)
    }
    
    def testReadCell_String() {
        when:
        Path excel = workdir_.resolve("testReadCell_String.xls")
        KSWorkbook wb = new KSWorkbook(excel)
        then:
        wb != null
        when:
        wb.writeCell("Sheet1", 0, 0, "Yes, yes, yes")
        String value = wb.readCell("Sheet1", 0, 0)
        then:
        value == "Yes, yes, yes"
        when:
        wb.close()
        then:
        Files.exists(excel)
    }
    
    @Test
    void testReadCell_Numeric() {
        when:
        Path excel = workdir_.resolve("testReadCell_Numeric.xls")
        KSWorkbook wb = new KSWorkbook(excel)
        then:
        wb != null
        when:
        wb.writeCell("Sheet1", 0, 0, 1234.56)
        String value = wb.readCell("Sheet1", 0, 0)
        then:
        value == "1234.56"
        when:
        wb.close()
        then:
        Files.exists(excel)
    }
    
    void testReadCell_Date() {
        when:
        Path excel = workdir_.resolve("testReadCell_Date.xls")
        KSWorkbook wb = new KSWorkbook(excel)
        then:
        wb != null
        when:
        LocalDateTime ldt = LocalDateTime.of(2018, 10, 5, 11, 34, 00)
        Date date = Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant());
        wb.writeCell("Sheet1", 0, 0, date)
        String value = wb.readCell("Sheet1", 0, 0)
        then:
        value == "2018/10/05 11:34"
        when:
        wb.close()
        then:
        Files.exists(excel)
    }
}
