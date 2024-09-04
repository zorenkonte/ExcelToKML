package com.zorenkonte.exceltokml

import android.content.Context
import android.net.Uri
import androidx.lifecycle.LiveData
import androidx.lifecycle.MutableLiveData
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.BufferedInputStream
import java.io.InputStream

class ExcelViewModel : ViewModel() {
    private val _progress = MutableLiveData<Int>()
    val progress: LiveData<Int> get() = _progress

    private val _data = MutableLiveData<List<List<String>>>()
    val data: LiveData<List<List<String>>> get() = _data

    fun readExcelFile(context: Context, uri: Uri) {
        viewModelScope.launch(Dispatchers.IO) {
            val contentResolver = context.contentResolver
            val inputStream: InputStream? = contentResolver.openInputStream(uri)
            val data = mutableListOf<List<String>>()
            val formatter = DataFormatter()

            inputStream.use { stream ->
                val bufferedStream = BufferedInputStream(stream)
                val workbook = WorkbookFactory.create(bufferedStream)
                val sheet = workbook.getSheetAt(0)
                val totalRows = sheet.physicalNumberOfRows
                for ((rowIndex, row) in sheet.withIndex()) {
                    val rowData = mutableListOf<String>()
                    for (cell in row) {
                        rowData.add(formatter.formatCellValue(cell))
                    }
                    data.add(rowData)
                    _progress.postValue((rowIndex + 1) * 100 / totalRows)
                }
            }

            _data.postValue(data)
        }
    }

    fun resetProgress() {
        _progress.value = 0
    }
}