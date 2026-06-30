@file:OptIn(ExperimentalMaterial3Api::class, ExperimentalMaterial3ExpressiveApi::class)

package com.zorenkonte.exceltokml.ui

import android.content.Context
import android.content.Intent
import android.net.Uri
import android.widget.Toast
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.animation.core.animateFloatAsState
import androidx.compose.animation.core.tween
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material3.Button
import androidx.compose.material3.CircularWavyProgressIndicator
import androidx.compose.material3.ElevatedCard
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.ExperimentalMaterial3ExpressiveApi
import androidx.compose.material3.FilledTonalButton
import androidx.compose.material3.LinearWavyProgressIndicator
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.zorenkonte.exceltokml.ConversionPhase
import com.zorenkonte.exceltokml.ExcelUiState
import com.zorenkonte.exceltokml.ExcelViewModel

private val ExcelMimeTypes = arrayOf(
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

@Composable
fun HomeScreen(viewModel: ExcelViewModel) {
    val context = LocalContext.current
    val uiState by viewModel.uiState.collectAsStateWithLifecycle()

    val filePickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        uri?.let { viewModel.readExcelFile(context, it) }
    }

    ConverterContent(
        uiState = uiState,
        onConvertClick = { filePickerLauncher.launch(ExcelMimeTypes) },
        onOpenKmlClick = { openKml(context, viewModel, uiState.rows) }
    )
}

@Composable
private fun ConverterContent(
    uiState: ExcelUiState,
    onConvertClick: () -> Unit,
    onOpenKmlClick: () -> Unit
) {
    val animatedProgress by animateFloatAsState(
        targetValue = uiState.progress / 100f,
        animationSpec = tween(durationMillis = 500),
        label = "progress"
    )

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        horizontalAlignment = Alignment.CenterHorizontally,
        verticalArrangement =
            if (uiState.rows.isEmpty()) Arrangement.Center else Arrangement.Top
    ) {
        if (uiState.phase == ConversionPhase.READING || uiState.phase == ConversionPhase.PARSING) {
            ProgressSection(phase = uiState.phase, animatedProgress = animatedProgress)
            Spacer(modifier = Modifier.height(24.dp))
        }

        if (uiState.phase == ConversionPhase.IDLE || uiState.phase == ConversionPhase.DONE) {
            ActionButtons(
                hasData = uiState.rows.isNotEmpty(),
                onConvertClick = onConvertClick,
                onOpenKmlClick = onOpenKmlClick
            )
        }

        if (uiState.rows.isNotEmpty()) {
            Spacer(modifier = Modifier.height(16.dp))
            LazyColumn(
                modifier = Modifier
                    .fillMaxWidth()
                    .weight(1f)
            ) {
                items(uiState.rows) { row -> PlaceRecordCard(row) }
            }
        }
    }
}

@Composable
private fun ProgressSection(phase: ConversionPhase, animatedProgress: Float) {
    Column(horizontalAlignment = Alignment.CenterHorizontally) {
        when (phase) {
            ConversionPhase.READING -> {
                // Expressive indeterminate spinner while the file is being opened.
                CircularWavyProgressIndicator()
                Spacer(modifier = Modifier.height(12.dp))
                Text(
                    text = "Reading file…",
                    style = MaterialTheme.typography.titleMedium
                )
            }

            ConversionPhase.PARSING -> {
                // Expressive determinate wavy bar tracks per-row parsing progress.
                LinearWavyProgressIndicator(
                    progress = { animatedProgress },
                    modifier = Modifier.fillMaxWidth()
                )
                Spacer(modifier = Modifier.height(12.dp))
                Text(
                    text = "${(animatedProgress * 100).toInt()}%",
                    style = MaterialTheme.typography.headlineMedium,
                    fontWeight = FontWeight.Medium
                )
            }

            else -> Unit
        }
    }
}

@Composable
private fun ActionButtons(
    hasData: Boolean,
    onConvertClick: () -> Unit,
    onOpenKmlClick: () -> Unit
) {
    Row(
        horizontalArrangement = Arrangement.spacedBy(12.dp),
        verticalAlignment = Alignment.CenterVertically
    ) {
        Button(onClick = onConvertClick) {
            Text(text = if (hasData) "Convert again" else "Convert")
        }
        if (hasData) {
            FilledTonalButton(onClick = onOpenKmlClick) {
                Text(text = "Open KML")
            }
        }
    }
}

@Composable
private fun PlaceRecordCard(row: List<String>) {
    fun cell(index: Int): String = row.getOrElse(index) { "" }

    ElevatedCard(
        modifier = Modifier
            .fillMaxWidth()
            .padding(vertical = 6.dp)
    ) {
        Column(modifier = Modifier.padding(16.dp)) {
            Text(
                text = "CAN: ${cell(0)}",
                style = MaterialTheme.typography.titleMedium,
                color = MaterialTheme.colorScheme.primary
            )
            Spacer(modifier = Modifier.height(4.dp))
            Text(text = "Name: ${cell(4)}", style = MaterialTheme.typography.bodyMedium)
            Text(text = "Address: ${cell(5)}", style = MaterialTheme.typography.bodyMedium)
            Text(text = "Street: ${cell(6)}", style = MaterialTheme.typography.bodyMedium)
            Text(
                text = "Location: ${cell(10)}, ${cell(9)}",
                style = MaterialTheme.typography.bodyMedium
            )
        }
    }
}

private fun openKml(context: Context, viewModel: ExcelViewModel, rows: List<List<String>>) {
    val kmlUri = viewModel.convertToKML(rows, context)
    if (kmlUri == null) {
        Toast.makeText(context, "Failed to create KML file", Toast.LENGTH_SHORT).show()
        return
    }

    val intent = Intent(Intent.ACTION_VIEW).apply {
        setDataAndType(kmlUri, "application/*")
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }
    val chooser = Intent.createChooser(intent, "Open KML file with")
    if (intent.resolveActivity(context.packageManager) != null) {
        context.startActivity(chooser)
    } else {
        Toast.makeText(context, "No application found to open KML file", Toast.LENGTH_SHORT).show()
    }
}
