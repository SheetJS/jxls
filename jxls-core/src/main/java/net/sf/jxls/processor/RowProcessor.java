package net.sf.jxls.processor;

import net.sf.jxls.transformer.Row;

import java.util.Map;

/**
 * Allows dynamic processing of rows
 * @author <a href="mailto:Lvissochin@db.luxoft.com">Leonid Vysochin</a>
 */
public interface RowProcessor {
    void processRow(Row row, Map namedCells);
}
