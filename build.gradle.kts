// Top-level build file where you can add configuration options common to all sub-projects/modules.
plugins {
    // AGP 9 provides built-in Kotlin support; the standalone kotlin.android plugin is gone.
    alias(libs.plugins.android.application) apply false
    alias(libs.plugins.kotlin.compose) apply false
}
