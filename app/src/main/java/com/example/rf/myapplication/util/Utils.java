package com.example.rf.myapplication.util;

import android.os.Environment;
import android.util.Log;

import java.io.File;

/**
 * Created by RF on 2018/5/2.
 */

public class Utils {
    public static String getExcelDir() {
        // SD卡指定文件夹
        String sdcardPath = Environment.getExternalStorageDirectory()
                .toString();
        File dir = new File(sdcardPath + File.separator + "Excel"
                + File.separator + "Person");
        if (!dir.exists()) {
            dir.mkdirs();
        }
        if (dir.exists()) {
            return dir.toString();

        } else {
            dir.mkdirs();
            Log.e("BAG", "保存路径不存在,");
            return dir.toString();
        }

    }
}
