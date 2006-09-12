package net.sf.jxls.util;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.beanutils.BeanPropertyValueEqualsPredicate;

import java.util.*;
import java.lang.reflect.InvocationTargetException;

/**
 * @author Leonid Vysochyn
 */
public class ReportUtil {

    public static Collection groupCollectionData(Collection objects, String groupBy) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        Collection result = new ArrayList();
        if( objects != null ){
            Set groupByValues = new TreeSet(); // using TreeSet to ensure groups are sorted according to natural order
            for (Iterator iterator = objects.iterator(); iterator.hasNext();) {
                Object bean = iterator.next();
                groupByValues.add( PropertyUtils.getProperty( bean, groupBy) );
            }
            for (Iterator iterator = groupByValues.iterator(); iterator.hasNext();) {
                Object groupValue = iterator.next();
                BeanPropertyValueEqualsPredicate predicate = new BeanPropertyValueEqualsPredicate( groupBy, groupValue );
                Collection groupItems = CollectionUtils.select( objects, predicate );
                GroupData groupData = new GroupData( CollectionUtils.get( groupItems, 0 ), groupItems );
                result.add( groupData );
            }
        }
        return result;
    }

}
