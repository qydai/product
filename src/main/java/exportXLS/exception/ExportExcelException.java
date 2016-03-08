package exportXLS.exception;

/**
 * Created by qydai on 2016/3/8.
 */
public class ExportExcelException extends Exception {
    public ExportExcelException(){

    }
    public ExportExcelException(String exceptionInfo){
        super(exceptionInfo);
    }
    public ExportExcelException(Exception e){
        super(e);
    }

}
