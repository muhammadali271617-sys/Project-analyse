package uz.f7.trainer

import android.annotation.SuppressLint
import android.graphics.Color
import android.os.Bundle
import android.webkit.WebChromeClient
import android.webkit.WebResourceRequest
import android.webkit.WebResourceResponse
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.FrameLayout
import androidx.activity.ComponentActivity
import androidx.activity.OnBackPressedCallback
import androidx.core.splashscreen.SplashScreen.Companion.installSplashScreen
import androidx.core.view.ViewCompat
import androidx.core.view.WindowCompat
import androidx.core.view.WindowInsetsCompat
import androidx.webkit.WebViewAssetLoader
import java.io.ByteArrayInputStream

class MainActivity : ComponentActivity() {

    private lateinit var webView: WebView

    private val assetLoader by lazy {
        WebViewAssetLoader.Builder()
            .setDomain(APP_HOST)
            .addPathHandler("/assets/", WebViewAssetLoader.AssetsPathHandler(this))
            .build()
    }

    private val backCallback = object : OnBackPressedCallback(false) {
        override fun handleOnBackPressed() {
            if (webView.canGoBack()) webView.goBack()
        }
    }

    @SuppressLint("SetJavaScriptEnabled")
    override fun onCreate(savedInstanceState: Bundle?) {
        installSplashScreen()
        super.onCreate(savedInstanceState)

        // Draw edge-to-edge and paint the system-bar areas with the brand colour
        // (window.statusBarColor is ignored on Android 15+).
        WindowCompat.setDecorFitsSystemWindows(window, false)
        @Suppress("DEPRECATION")
        window.statusBarColor = BRAND_COLOR
        WindowCompat.getInsetsController(window, window.decorView).apply {
            isAppearanceLightStatusBars = false
            isAppearanceLightNavigationBars = false
        }

        val root = FrameLayout(this).apply { setBackgroundColor(BRAND_COLOR) }
        webView = WebView(this).apply { setBackgroundColor(Color.WHITE) }
        root.addView(
            webView,
            FrameLayout.LayoutParams(
                FrameLayout.LayoutParams.MATCH_PARENT,
                FrameLayout.LayoutParams.MATCH_PARENT
            )
        )
        ViewCompat.setOnApplyWindowInsetsListener(root) { v, insets ->
            val bars = insets.getInsets(
                WindowInsetsCompat.Type.systemBars() or
                    WindowInsetsCompat.Type.displayCutout() or
                    WindowInsetsCompat.Type.ime()
            )
            v.setPadding(bars.left, bars.top, bars.right, bars.bottom)
            WindowInsetsCompat.CONSUMED
        }
        setContentView(root)

        webView.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true
            databaseEnabled = true
            allowFileAccess = false
            allowContentAccess = false
            cacheMode = WebSettings.LOAD_NO_CACHE
            textZoom = 100
            setSupportZoom(false)
            builtInZoomControls = false
            displayZoomControls = false
            loadWithOverviewMode = true
            useWideViewPort = true
            setSupportMultipleWindows(false)
            javaScriptCanOpenWindowsAutomatically = false
        }

        // Needed so that alert()/confirm()/prompt() dialogs work.
        webView.webChromeClient = WebChromeClient()

        webView.webViewClient = object : WebViewClient() {
            override fun shouldInterceptRequest(
                view: WebView,
                request: WebResourceRequest
            ): WebResourceResponse? {
                val url = request.url
                if (url.host == APP_HOST) return assetLoader.shouldInterceptRequest(url)
                // Anything else is blocked: the app is fully offline.
                return WebResourceResponse(
                    "text/plain", "utf-8", 403, "Blocked", emptyMap(),
                    ByteArrayInputStream(ByteArray(0))
                )
            }

            override fun shouldOverrideUrlLoading(
                view: WebView,
                request: WebResourceRequest
            ): Boolean {
                // Allow only our own local pages; never leave the app.
                return request.url.host != APP_HOST
            }

            override fun doUpdateVisitedHistory(view: WebView, url: String?, isReload: Boolean) {
                super.doUpdateVisitedHistory(view, url, isReload)
                backCallback.isEnabled = view.canGoBack()
            }
        }

        onBackPressedDispatcher.addCallback(this, backCallback)

        val restored = savedInstanceState?.let { webView.restoreState(it) } != null
        if (!restored) webView.loadUrl(START_URL)
        backCallback.isEnabled = webView.canGoBack()
    }

    override fun onSaveInstanceState(outState: Bundle) {
        super.onSaveInstanceState(outState)
        webView.saveState(outState)
    }

    override fun onPause() {
        webView.onPause()
        super.onPause()
    }

    override fun onResume() {
        super.onResume()
        webView.onResume()
    }

    override fun onDestroy() {
        (webView.parent as? FrameLayout)?.removeView(webView)
        webView.destroy()
        super.onDestroy()
    }

    private companion object {
        const val APP_HOST = "appassets.androidplatform.net"
        const val START_URL = "https://$APP_HOST/assets/index.html"
        val BRAND_COLOR = Color.parseColor("#3730A3")
    }
}
