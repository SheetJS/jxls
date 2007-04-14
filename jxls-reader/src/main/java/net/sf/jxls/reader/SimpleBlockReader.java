package net.sf.jxls.reader;

import java.util.List;

/**
 * Interface to read simple block of excel rows
 * @author Leonid Vysochyn
 */
public interface SimpleBlockReader extends XLSBlockReader{
    void addMapping(BeanCellMapping mapping);

    List getMappings();
}
