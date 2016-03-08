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
	public String name() default "";
	public int sortNo() default 0;
	public String formatter() default ""; 
}