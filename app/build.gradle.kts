plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
}

val runNumber: Int = System.getenv("GITHUB_RUN_NUMBER")?.toIntOrNull() ?: 1

android {
    namespace = "uz.f7.trainer"
    compileSdk = 36

    defaultConfig {
        applicationId = "uz.f7.trainer"
        minSdk = 24
        targetSdk = 36
        versionCode = runNumber
        versionName = "1.0.$runNumber"
    }

    buildTypes {
        release {
            isMinifyEnabled = false
        }
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    // HTML must be packed as-is (no compression issues for large single file)
    androidResources {
        noCompress += "html"
    }
}

kotlin {
    jvmToolchain(17)
}

dependencies {
    implementation("androidx.core:core-ktx:1.16.0")
    implementation("androidx.activity:activity-ktx:1.10.1")
    implementation("androidx.webkit:webkit:1.14.0")
    implementation("androidx.core:core-splashscreen:1.0.1")
}
