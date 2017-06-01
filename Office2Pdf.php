/**
 * base on LibreOffice
 *
 * @author gz.tony@foxmail.com
 */
class Office2Pdf
{
    protected static $_file2pdf =  [];

    public static function isSupportDocType($ext)
    {
       $extArr =  ['doc','xls','ppt','docx','xlsx','pptx','wps','wpt','dot','rtf','dps','dpt','pot','pps','et','ett','xlt'];
       return in_array(strval($ext),$extArr);
    }   
    
    public static function convert($doc,$path)
    {
        $array = explode('.', $doc);
        $ext = array_pop($array);
        if (!self::isSupportDocType($ext) || !is_dir($path)) {
            return false;
        }
        $pdf = implode('.', $array).DIRECTORY_SEPARATOR.'.pdf';
        if (self::been2Pdf($pdf,$path)) {
            return true;
        }
        self::$_file2pdf[] = $pdf;
        if (strtolower(substr(PHP_OS, 0, 3)) == 'win') {
            pclose(popen( "start /B " .'soffice --invisible --convert-to pdf:writer_pdf_Export --outdir "'.$path.'" "'.$doc.'"' , "r"));
        } else {
            shell_exec('soffice --headless --convert-to pdf:writer_pdf_Export --outdir '.$path.' '.$doc.'  > /dev/null &');
        }
        return true;
    }

    public static function been2Pdf($pdf,$path)
    {
        if (in_array($pdf, self::$_file2pdf)) {
            return true;
        }
        if (is_file($path.DIRECTORY_SEPARATOR.$pdf)) {
            return true;
        }
        if (is_file($path.DIRECTORY_SEPARATOR.'.~lock.'.$pdf.'#')) {
            return true;
        }
        return false;
    }
    
}