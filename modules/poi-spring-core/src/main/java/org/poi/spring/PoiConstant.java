package org.poi.spring;

/**
 * Created by Hong.LvHang on 2017-04-13.
 */
public abstract class PoiConstant {
    private PoiConstant() {

    }

    public static final String EMPTY_STRING = "";
    public static final String CLASS_NAME_SUFFIX = ".EXCLE";
    public static final int DEFAULT_WIDTH_R = 1024;
    public static final int DEFAULT_WIDTH = 4;

    public static final String TRUE_VALUE = "true";
    public static final String FALSE_VALUE = "false";


    public static final class ExcleDefinitionProperties {
        public static final String DATA_CLASS = "dataClass";
        public static final String ID = "id";
        public static final String EXCLE_NAME = "excleName";
        public static final String SHEET_NAME = "sheetName";
        public static final String SHEET_INDEX = "sheetIndex";
        public static final String HEADER = "header";
        public static final String COLUMN_WIDTH = "columnWidth";
        public static final String DEFAULT_PROPERTIES = "defaultProperties";
        public static final String COLUMN_DEFINITIONS = "columnDefinitions";
    }

    public static final class ExcleAnnProperties {
        public static final String EXCLE_NAME = "name";
        public static final String SHEET_NAME = "sheetName";
        public static final String SHEET_INDEX = "sheetIndex";
        public static final String HEADER = "header";
        public static final String WIDTH = "width";
        public static final String TITLE = "title";
        public static final String REGEX = "regex";
        public static final String REQUIRED = "required";
        public static final String FORMAT = "format";
        public static final String DICTNO = "dictNo";
        public static final String FGCOLOR = "fgcolor";
        public static final String BORDER = "border";
        public static final String ALIGN = "align";
    }


    public static final class ExcleXmlProperties {
        public static final String EXCLE_NAME = "excle-name";

    }


    public static final String DATA_CLASS_ATTRIBUTE = "data-class";
    public static final String EXCLE_NAME_ATTRIBUTE = "excle-name";
    public static final String SHEET_NAME_ATTRIBUTE = "sheet-name";
    public static final String SHEET_INDEX_ATTRIBUTE = "sheet-index";
    public static final String DEFAULT_COLUMN_WIDTH_ATTRIBUTE = "column-width";
    public static final String DEFAULT_ALIGN_ATTRIBUTE = "default-align";

    //COLUMN
    public static final String COLUMN_ELEMENT = "poi:column";
    public static final String COLUMN_NAME_ATTRIBUTE = "name";
    public static final String TITLE_ATTRIBUTE = "title";
    public static final String REGEX_ATTRIBUTE = "regex";
    public static final String REQUIRED_ATTRIBUTE = "required";
    public static final String ALIGN_ATTRIBUTE = "align";
    public static final String COLUMN_WIDTH_ATTRIBUTE = "column-width";
    public static final String DEFAULT_VALUE_ATTRIBUTE = "default-value";


}
