package com.marzaha.pptx;

import android.Manifest;
import android.app.Activity;
import android.app.ProgressDialog;
import android.content.pm.PackageManager;
import android.os.AsyncTask;
import android.os.Bundle;
import android.os.Handler;
import android.support.annotation.NonNull;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.webkit.WebSettings;
import android.webkit.WebView;

import java.io.File;


public class PptxActivity extends Activity {

    private WebView webView;
    private String filePath = null;
    private Handler handler;

    private ProgressDialog progressDialog;

    @Override
    protected void onCreate(Bundle savedInstanceState) {

        System.setProperty("javax.xml.stream.XMLInputFactory", "com.sun.xml.stream.ZephyrParserFactory");
        System.setProperty("javax.xml.stream.XMLOutputFactory", "com.sun.xml.stream.ZephyrWriterFactory");
        System.setProperty("javax.xml.stream.XMLEventFactory", "com.sun.xml.stream.events.ZephyrEventFactory");
        Thread.currentThread().setContextClassLoader(getClass().getClassLoader());

        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_pptx);

        init();

    }

    private void init() {
        webView = findViewById(R.id.webview);
        webView.getSettings().setJavaScriptEnabled(true);
        webView.getSettings().setSupportZoom(true);
        webView.getSettings().setBuiltInZoomControls(true);
        webView.getSettings().setDisplayZoomControls(false);
        webView.getSettings().setUseWideViewPort(true);
        webView.getSettings().setLoadWithOverviewMode(true);
        webView.getSettings().setTextZoom(200);
        webView.getSettings().setLayoutAlgorithm(WebSettings.LayoutAlgorithm.SINGLE_COLUMN);


        filePath = "/sdcard/Download/test.pptx";
        progressDialog = new ProgressDialog(this);

        if (new File(filePath).canRead()) {
            openFile();
        } else {
            String permission = Manifest.permission.READ_EXTERNAL_STORAGE;
            int check = ContextCompat.checkSelfPermission(this, permission);
            if (check != PackageManager.PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(this, new String[]{permission}, 1);
            }
        }

    }

    public class ConvertTask extends AsyncTask<String, Void, String> {
        @Override
        protected String doInBackground(String... objects) {
            String filePath = objects[0];
            PointParser pointParser = new PointParser(PptxActivity.this, filePath);
            pointParser.pptxToHtml(handler);
            return pointParser.htmlPath;
        }

        @Override
        protected void onPostExecute(String returnString) {
            progressDialog.dismiss();
            webView.loadUrl("file:///" + returnString);
        }
    }

    private void openFile() {
        handler = new Handler();
        progressDialog.setMessage("正在转换文件....");
        progressDialog.show();

        new ConvertTask().execute(filePath);
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        if (grantResults[0] == PackageManager.PERMISSION_GRANTED) {
            openFile();
        } else {
            finish();
        }
    }
    

}