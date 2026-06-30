@file:OptIn(ExperimentalMaterial3Api::class)

package com.zorenkonte.exceltokml.ui

import androidx.compose.foundation.layout.padding
import androidx.compose.material3.DrawerState
import androidx.compose.material3.DrawerValue
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.ModalDrawerSheet
import androidx.compose.material3.ModalNavigationDrawer
import androidx.compose.material3.NavigationDrawerItem
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.material3.TopAppBar
import androidx.compose.material3.rememberDrawerState
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.rememberCoroutineScope
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.SolidColor
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.res.stringResource
import androidx.compose.ui.unit.dp
import androidx.navigation.NavHostController
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.currentBackStackEntryAsState
import androidx.navigation.compose.rememberNavController
import com.zorenkonte.exceltokml.ExcelViewModel
import com.zorenkonte.exceltokml.R
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.launch

private enum class Destination(
    val route: String,
    val label: String
) {
    HOME("home", "Home"),
    CREDITS("credits", "Credits")
}

@Composable
fun ExcelToKmlApp(viewModel: ExcelViewModel) {
    val navController = rememberNavController()
    val drawerState = rememberDrawerState(DrawerValue.Closed)
    val scope = rememberCoroutineScope()

    ModalNavigationDrawer(
        drawerState = drawerState,
        drawerContent = {
            AppDrawer(
                navController = navController,
                drawerState = drawerState,
                scope = scope
            )
        }
    ) {
        Scaffold(
            topBar = {
                TopAppBar(
                    title = { Text(stringResource(R.string.app_name)) },
                    navigationIcon = {
                        IconButton(onClick = { scope.launch { drawerState.open() } }) {
                            Icon(
                                imageVector = MenuIcon,
                                contentDescription = "Open navigation drawer"
                            )
                        }
                    }
                )
            }
        ) { innerPadding ->
            NavHost(
                navController = navController,
                startDestination = Destination.HOME.route,
                modifier = Modifier.padding(innerPadding)
            ) {
                composable(Destination.HOME.route) { HomeScreen(viewModel) }
                composable(Destination.CREDITS.route) { CreditsScreen() }
            }
        }
    }
}

@Composable
private fun AppDrawer(
    navController: NavHostController,
    drawerState: DrawerState,
    scope: CoroutineScope
) {
    val navBackStackEntry by navController.currentBackStackEntryAsState()
    val currentRoute = navBackStackEntry?.destination?.route

    ModalDrawerSheet {
        Destination.entries.forEach { destination ->
            NavigationDrawerItem(
                label = { Text(destination.label) },
                selected = currentRoute == destination.route,
                onClick = {
                    scope.launch { drawerState.close() }
                    if (currentRoute != destination.route) {
                        navController.navigate(destination.route) {
                            // Avoid stacking duplicate destinations and preserve state.
                            popUpTo(Destination.HOME.route) { saveState = true }
                            launchSingleTop = true
                            restoreState = true
                        }
                    }
                },
                modifier = Modifier.padding(horizontal = 12.dp)
            )
        }
    }
}

// Hamburger ("menu") icon defined inline so the app does not depend on the deprecated
// androidx.compose.material:material-icons artifacts, which are no longer a transitive
// dependency of material3 on the Compose 1.12 line. The fill color is irrelevant because
// Icon tints the vector with the current content color.
private val MenuIcon: ImageVector =
    ImageVector.Builder(
        name = "Menu",
        defaultWidth = 24.dp,
        defaultHeight = 24.dp,
        viewportWidth = 24f,
        viewportHeight = 24f
    ).apply {
        listOf(5f, 11f, 17f).forEach { top ->
            path(fill = SolidColor(Color.White)) {
                moveTo(3f, top)
                lineTo(21f, top)
                lineTo(21f, top + 2f)
                lineTo(3f, top + 2f)
                close()
            }
        }
    }.build()
