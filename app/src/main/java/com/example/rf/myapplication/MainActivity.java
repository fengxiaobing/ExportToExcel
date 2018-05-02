package com.example.rf.myapplication;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;

import com.example.rf.myapplication.util.SaveToExcelUtil;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;

import static com.example.rf.myapplication.util.Utils.getExcelDir;

public class MainActivity extends AppCompatActivity {
    String excelPath = getExcelDir()+ File.separator+"demo.xls";
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        findViewById(R.id.tv).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SaveToExcelUtil saveToExcel = new SaveToExcelUtil(MainActivity.this,excelPath);
                saveToExcel.writeToExcel(22,33,44,55);
            }
        });

    }


}
