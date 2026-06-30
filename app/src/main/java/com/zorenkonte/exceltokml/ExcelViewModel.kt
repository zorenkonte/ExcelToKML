package com.zorenkonte.exceltokml

import android.content.Context
import android.net.Uri
import androidx.core.content.FileProvider
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.flow.update
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.BufferedInputStream
import java.io.BufferedWriter
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.FileWriter
import java.io.InputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/** The stage of the Excel -> KML conversion flow, used to drive the UI. */
enum class ConversionPhase {
    /** Nothing is happening; waiting for the user to pick a file. */
    IDLE,

    /** A file was picked and is being opened (work has not started reporting progress yet). */
    READING,

    /** Rows are actively being parsed; [ExcelUiState.progress] is meaningful (1..100). */
    PARSING,

    /** Parsing finished; [ExcelUiState.rows] holds the result. */
    DONE
}

/** Immutable snapshot of everything the home screen needs to render. */
data class ExcelUiState(
    val phase: ConversionPhase = ConversionPhase.IDLE,
    val progress: Int = 0,
    val rows: List<List<String>> = emptyList()
)

class ExcelViewModel : ViewModel() {
    private val _uiState = MutableStateFlow(ExcelUiState())
    val uiState: StateFlow<ExcelUiState> = _uiState.asStateFlow()

    fun readExcelFile(context: Context, uri: Uri) {
        _uiState.value = ExcelUiState(phase = ConversionPhase.READING)
        viewModelScope.launch(Dispatchers.IO) {
            try {
                val inputStream: InputStream? = context.contentResolver.openInputStream(uri)
                val formatter = DataFormatter() // Format cell values (including empty cells)
                val rows = mutableListOf<List<String>>()

                inputStream.use { stream ->
                    val bufferedStream = BufferedInputStream(stream)
                    val workbook = WorkbookFactory.create(bufferedStream)
                    val sheet = workbook.getSheetAt(0)
                    // Guard against an empty sheet so progress math never divides by zero.
                    val totalRows = (sheet.physicalNumberOfRows - 1).coerceAtLeast(1)

                    // Iterate over each row
                    sheet.drop(1).forEachIndexed { rowIndex, row ->
                        val rowData = mutableListOf<String>()
                        // Ensure that we process all cells, including empty ones
                        for (cellIndex in 0 until row.lastCellNum) {
                            val cell = row.getCell(cellIndex) // Get cell at this index
                            val cellValue = formatter.formatCellValue(cell) // Format the cell value
                            rowData.add(cellValue) // Add the value (even if it's empty)
                        }
                        rows.add(rowData)
                        updateProgress(rowIndex, totalRows)
                    }
                }

                _uiState.value = ExcelUiState(
                    phase = ConversionPhase.DONE,
                    progress = 100,
                    rows = rows
                )
            } catch (e: Exception) {
                e.printStackTrace()
                // Reset to a clean state so the user can try another file.
                _uiState.value = ExcelUiState()
            }
        }
    }

    private fun updateProgress(rowIndex: Int, totalRows: Int) {
        val percent = ((rowIndex + 1) * 100 / totalRows).coerceIn(0, 100)
        _uiState.update { it.copy(phase = ConversionPhase.PARSING, progress = percent) }
    }

    fun convertToKML(data: List<List<String>>, context: Context): Uri? {
        val kmlFile = writeKMLToStream(data, context) // Write KML progressively
        return kmlFile?.let { compressToKMZ(it, context) } // Compress to KMZ after writing
    }

    private fun writeKMLToStream(data: List<List<String>>, context: Context): File? {
        return try {
            val file = File(context.cacheDir, "output.kml")
            BufferedWriter(FileWriter(file)).use { writer ->
                writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
                writer.write("<kml xmlns=\"http://www.opengis.net/kml/2.2\">\n")
                writer.write("<Document>\n")

                data.drop(1).forEach { row ->
                    writer.write(createPlaceMark(row)) // Write each placemark in chunks
                }

                writer.write("</Document>\n")
                writer.write("</kml>\n")
            }
            file
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }

    private fun compressToKMZ(kmlFile: File, context: Context): Uri? {
        return try {
            val kmzFile = File(context.cacheDir, "output.kmz")
            ZipOutputStream(FileOutputStream(kmzFile)).use { zos ->
                zos.putNextEntry(ZipEntry(kmlFile.name)) // Add KML to KMZ
                FileInputStream(kmlFile).use { fis ->
                    fis.copyTo(zos) // Stream KML content to KMZ
                }
                zos.closeEntry() // Finish entry
            }
            FileProvider.getUriForFile(context, "${context.packageName}.fileprovider", kmzFile)
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }

    private fun createPlaceMark(row: List<String>): String {
        // Tolerate rows that are shorter than expected instead of crashing.
        fun cell(index: Int): String = row.getOrElse(index) { "" }
        return """
            <Placemark>
                <name>${cell(0)}</name>
                <description><![CDATA[
                    <p>CAN: ${cell(0)}<br/>
                       Meter Code: ${cell(1)}<br/>
                       MRU: ${cell(2)}<br/>
                       BA: ${cell(3)}<br/>
                       Name: ${cell(4)}<br/>
                       Address: ${cell(5)}<br/>
                       DMA: ${cell(6)}<br/>
                       DMZ: ${cell(7)}<br/>
                       Latitude: ${cell(9)}<br/>
                       Longitude: ${cell(10)}<br/>
                       Coordinates Source: ${cell(11)}</p>
                ]]></description>
                <Point><coordinates>${cell(10)},${cell(9)}</coordinates></Point>
            </Placemark>
        """.trimIndent()
    }
}
