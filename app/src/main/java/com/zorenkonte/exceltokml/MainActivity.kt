package com.zorenkonte.exceltokml

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.activity.viewModels
import com.zorenkonte.exceltokml.ui.ExcelToKmlApp
import com.zorenkonte.exceltokml.ui.theme.ExcelToKMLTheme

class MainActivity : ComponentActivity() {
    private val excelViewModel: ExcelViewModel by viewModels()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContent {
            ExcelToKMLTheme {
                ExcelToKmlApp(excelViewModel)
            }
        }
    }
}
