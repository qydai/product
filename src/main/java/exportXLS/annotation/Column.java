package exportXLS.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Column {
    /**
     * 列名
     * @return
     */
    public String name() default "";

    /**
     * 排序字段，不写为默认排序
     * @return
     */
    public int sortNo() default 0;

    /**
     * 定义本类中一个格式化方法，参数为字段本身
     * @return
     */
    public String formatter() default "";

    /**
     * 只支持jpge,png两种图片的导出
     * @return jpg or png
     */
    public String imgType()default "";
}
