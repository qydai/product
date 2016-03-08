package exportXLS.annotation;

import java.lang.annotation.*;
/**
 * @author JiaoYuQing
 *
 */
@Target(ElementType.TYPE)			
@Retention(RetentionPolicy.RUNTIME) 
@Documented 						
@Inherited							
public @interface Xls {
	public String name() default "";
}
