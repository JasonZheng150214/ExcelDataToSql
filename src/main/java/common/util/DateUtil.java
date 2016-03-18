package common.util;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by jason on 16/3/18.
 */
public class DateUtil {


    public static String convertDatetYMDHMS(Date date) {
        if (date == null) {
            return null;
        }
        SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-DD HH:mm:ss");
        return sf.format(date);
    }

    public static String convertDatetYMD(Date date) {
        if (date == null) {
            return null;
        }
        SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-DD");
        return sf.format(date);
    }
}
