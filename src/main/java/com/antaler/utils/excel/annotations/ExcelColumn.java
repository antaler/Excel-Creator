package com.antaler.utils.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.time.LocalDateTime;

import com.antaler.utils.excel.enums.ExcelDateFormat;
/**
 * Represent a column in excel
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelColumn {
    /**
     * order in column from left to right. Mandatory
     * @return  order in column
     */
    int order();
    /**
     * title from column. Mandatory
     * @return title from column
     */
    String title();

    /**
     * Text in the sheet cell for {@code true} value . Default Value {@code "True"}
     * @return text if value was {@code true}
     */
    String trueValue() default "True";

     /**
     * Text in the sheet cell for {@code false} value . Default Value {@code "False"}
     * @return text if value was {@code false}
     */
    String falseValue() default "False";

    /**
     * Set format for date if the  field type is {@link LocalDateTime}. Default Value is {@link ExcelDateFormat#NUMBER_SHORT_WITH_TIME}
     * @return format enum
     */
    ExcelDateFormat dateFormat() default ExcelDateFormat.NUMBER_SHORT_WITH_TIME;
}
