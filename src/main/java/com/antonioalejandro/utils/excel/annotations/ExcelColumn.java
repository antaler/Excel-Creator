package com.antonioalejandro.utils.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.antonioalejandro.utils.excel.enums.ExcelDateFormat;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelColumn {
    int order();
    String title();

    String trueValue() default "True";
    String falseValue() default "False";
    ExcelDateFormat dateFormat() default ExcelDateFormat.NUMBER_SHORT_WITH_TIME;
}
