package net.sf.jxls.controller;

import org.apache.poi.hssf.util.CellReference;
import net.sf.jxls.tag.Point;
import net.sf.jxls.tag.Block;
import net.sf.jxls.transformer.RowCollection;

import java.util.*;

import net.sf.jxls.controller.MatrixCell;

/**
 * This class tracks all cells transformations using virtual object matrix
 * @author Leonid Vysochyn
 */
public class TransformationMatrix {
    List matrix;
    Map cells = new HashMap();
    Map cellsMapping = new HashMap();
    Map duplicatedFormulaCellsMapping = new HashMap();
    Map duplicatedFormulaCells = new HashMap();

    boolean isCellMappingsReady;


    public TransformationMatrix(int rowNum, int colNum) {
        init( rowNum, colNum );
    }

    public int getRowCount(){
        return matrix.size();
    }

    public int getColCount(){
        return ((List)matrix.get(0)).size();
    }

    public void buildCellMappings(){
        Set sourceCells = cells.keySet();
        for (Iterator iterator = sourceCells.iterator(); iterator.hasNext();) {
            MatrixCell cell = (MatrixCell) iterator.next();
            Object mappedObject = cells.get( cell );
            cellsMapping.put( cell, findObjectCells( mappedObject ) );
        }
        isCellMappingsReady = true;
    }

    public List findMappedCells(int rowNum, int colNum, boolean useBuiltCellMappings){
        MatrixCell cell = new MatrixCell(rowNum, colNum);
        if( useBuiltCellMappings ){
            if( !isCellMappingsReady ){
                buildCellMappings();
            }
            return (List) cellsMapping.get( cell );
        }else{
            return findMappedCells( rowNum, colNum );
        }
    }


    public List findMappedCells(int rowNum, int colNum){
        List mappedCells = new ArrayList();
        MatrixCell cell = new MatrixCell(rowNum, colNum);
        Object mappedObject = cells.get( cell );
        if( mappedObject != null ){
            mappedCells = findObjectCells( mappedObject );
        }
        return mappedCells;
    }

    private List findObjectCells(Object obj) {
        List cells = new ArrayList();
        for(int i = 0; i < getRowCount(); i++){
            List row = (List) matrix.get( i );
            for (int j = 0; j < row.size(); j++) {
                Object item = row.get(j);
                if( item == obj ){
                    cells.add( new MatrixCell( i, j ) );
                }
            }
        }
        return cells;
    }

    public void init(int rowNum, int colNum){
        int dim = rowNum * colNum;
        matrix = new ArrayList( dim );
        matrix.addAll( Collections.nCopies( rowNum, null) );
        for(int i = 0; i < rowNum; i++){
            matrix.set(i, new ArrayList( dim ));
        }

        for(int i = 0; i < rowNum; i++){
            List row = (List) matrix.get(i);
            row.addAll( Collections.nCopies( colNum, null ) );
            for(int j = 0; j < colNum; j++){
                Object obj = new Object();
                row.set( j, obj );
                cells.put( new MatrixCell( i, j ), obj );
            }
        }
    }

    public void shift(int startRowNum, int startColNum, int endColNum, int n){
        if( n < 0 ){
            // shift up
            for(int i = startRowNum; i < getRowCount(); i++){
                List srcRow = (List) matrix.get( i );
                List destRow = (List) matrix.get( i + n );
                copyRange( destRow, srcRow, startColNum, endColNum);
                Collections.fill( srcRow.subList( startColNum, endColNum + 1), new Object());
            }
        }else if( n > 0 ){
            // shift down
            int destRowNum = startRowNum + n;
            matrix.addAll( Collections.nCopies( n, new ArrayList() ));
            for(int i = getRowCount() - 1; i >= destRowNum; i--){
                List srcRow = (List) matrix.get( i - n );
                List destRow = (List) matrix.get( i );
                if( destRow.size() < srcRow.size() ){
                    destRow = new ArrayList();
                    destRow.addAll( Collections.nCopies( getColCount(), new Object() ) );
                }
                copyRange( destRow, srcRow, startColNum, endColNum );
                matrix.set( i, destRow );
                Collections.fill( srcRow.subList( startColNum, endColNum + 1 ), new Object());
            }
        }
    }

    public void shiftColumns(int startRowNum, int endRowNum, int startColNum, int n){
        if( n < 0 ){
            // shift left
            for(int i = startRowNum; i <= endRowNum; i++){
                List row = (List) matrix.get( i );
                List newRow = new ArrayList();
                newRow.addAll( row.subList(0, startColNum + n) );
                newRow.addAll( row.subList( startColNum, row.size() ) );
                matrix.set( i, newRow );
            }
        }else if( n > 0 ){
            // shift right
            for(int i = startRowNum; i <= endRowNum; i++){
                List row = (List) matrix.get( i );
                row.addAll( startColNum, Collections.nCopies(n, new Object() ) );
            }
        }
    }


    private void copyRange(List destRow, List srcRow, int startColNum, int endColNum) {
        Collections.copy( destRow.subList( startColNum, endColNum + 1), srcRow.subList( startColNum, endColNum + 1) );
    }

    public void duplicateRows(int startRowNum, int endRowNum, int n) {
        int numberOfRows = (endRowNum - startRowNum + 1);
        int shiftNumber = numberOfRows * n;
        matrix.addAll( endRowNum + 1, Collections.nCopies( shiftNumber, new Object()) );
        for(int i = 0; i < n; i++){
            for(int j = 0; j < numberOfRows; j++){
                List srcRow = (List) matrix.get( startRowNum + j );
                List destRow = new ArrayList();
                destRow.addAll( srcRow );
                matrix.set( endRowNum + i * numberOfRows + j + 1, destRow );
            }
        }
    }

    public void duplicateRows(int startRowNum, int endRowNum, int n, Map formulaCellsToUpdate) {
        int numberOfRows = (endRowNum - startRowNum + 1);
        int shiftNumber = numberOfRows * n;
        matrix.addAll( endRowNum + 1, Collections.nCopies( shiftNumber, new Object()) );
        for(int i = 0; i < n; i++){
            for(int j = 0; j < numberOfRows; j++){
                List srcRow = (List) matrix.get( startRowNum + j );
                List destRow = new ArrayList();
                destRow.addAll( srcRow );
                matrix.set( endRowNum + i * numberOfRows + j + 1, destRow );
            }
            for (Iterator iterator = formulaCellsToUpdate.keySet().iterator(); iterator.hasNext();) {
                Point point = (Point) iterator.next();
                updateFormulaCells( numberOfRows * (i + 1), (List) formulaCellsToUpdate.get(point) );
            }
        }
    }

    private void updateFormulaCells(int shiftNumber, List refCellsToUpdate) {
        for (int i = 0; i < refCellsToUpdate.size(); i++) {
            String refCell = (String) refCellsToUpdate.get(i);
            CellReference cellReference = new CellReference( refCell );
            int rowNum = cellReference.getRow();
            rowNum += shiftNumber;
            int colNum = cellReference.getCol();
            MatrixCell matrixCell = new MatrixCell(rowNum, colNum);
            if( !cells.containsKey( matrixCell ) ){
                Object obj = new Object();
                set(rowNum, colNum, obj);
                cells.put( matrixCell, obj );
            }
        }
    }

    public void duplicate(int startRowNum, int endRowNum, int startColNum, int endColNum, int n, boolean shiftRows){
        int numberOfRows = endRowNum - startRowNum + 1;
        if( shiftRows ){
            shift( endRowNum + 1, startColNum, endColNum, numberOfRows * n );
        }
        for(int i = 0; i < n; i++){
            int destRowNum = endRowNum + 1 + numberOfRows * i;
            for(int j = startRowNum; j <= endRowNum; j++, destRowNum++){
                List srcRow = (List) matrix.get( j );
                List destRow;
                if( destRowNum >= matrix.size() ){
                    destRow = new ArrayList();
                    destRow.addAll( Collections.nCopies( srcRow.size(), new Object() ) );
                    matrix.add( destRow );
                }
                destRow = (List) matrix.get( destRowNum );
                if( destRow!=null && destRow.isEmpty() ){
                    destRow.addAll( Collections.nCopies( srcRow.size(), new Object() ) );
                }
                copyRange( destRow, srcRow, startColNum, endColNum );
            }
        }
    }

    public void shiftRows(int startRowNum, int n){
        if( n < 0 ){
            List newMatrix = new ArrayList();
            newMatrix.addAll( matrix.subList(0, startRowNum + n) );
            newMatrix.addAll( matrix.subList(startRowNum, matrix.size() ) );
            matrix = newMatrix;
        }else if( n > 0 ){
            matrix.addAll( startRowNum, Collections.nCopies( n, null ) );
            List newRow;
            for(int i = startRowNum; i < startRowNum + n; i++){
                newRow = new ArrayList();
                newRow.addAll( Collections.nCopies( getColCount(), new Object()));
                matrix.set( i, newRow );
            }
        }
    }


    public Object get(int i, int j){
        return ((List)matrix.get(i)).get(j);
    }

    public void set(int i, int j, Object obj){
        List row = (List) matrix.get( i );
        row.set( j, obj );
    }

    void removeRow(int i){
        matrix.remove( i );
    }

    public void removeBorders(int startRowNum, int endRowNum, int startColNum, int endColNum){
        shift( startRowNum + 1, startColNum, endColNum, -1 );
        shift( endRowNum, startColNum, endColNum, -1);
        shiftColumns( startRowNum, endRowNum - 2, startColNum + 1, -1 );
        shiftColumns( startRowNum, endRowNum - 2, endColNum, -1 );
    }

    public void removeBlockRows(Block block){

    }

    public void duplicateRow(RowCollection rowCollection){

    }

    public void setRow(int i, List newRow){
        List row = new ArrayList();
        row.addAll( Collections.nCopies( newRow.size(), null) );
        Collections.copy( row, newRow );
        matrix.set(i, row);
    }

    public void clearRow(int rowNum, int startColNum, int endColNum){
        List row = (List) matrix.get( rowNum );
        for (int i = startColNum; i <= endColNum; i++) {
            row.set( i, new Object() );
        }
    }

    public void removeLeftRightBorders(int startRowNum, int endRowNum, int startColNum, int endColNum){
        for(int i = startRowNum; i <= endRowNum; i++){
            List row = (List) matrix.get( i );
            row.remove( startColNum );
            row.remove( endColNum );
        }
    }

    public void duplicateRight(int startRowNum, int endRowNum, int startColNum, int endColNum, int n){
        int numberOfColumns = endColNum - startColNum + 1;
        int shiftNumber = numberOfColumns * n;
        for(int i = startRowNum; i <= endRowNum; i++){
            List row = (List) matrix.get( i );
            row.addAll( endColNum + 1, Collections.nCopies(shiftNumber, new Object()));
            for(int j = 0; j < n; j++){
                for(int k = startColNum; k <= endColNum; k++){
                    row.set( endColNum + numberOfColumns * j + k - startColNum + 1, row.get(k) );
                }
            }
        }
    }

    public Object clone() {
        TransformationMatrix newMatrix = new TransformationMatrix( getRowCount(), getColCount());
        for( int i = 0; i < getRowCount(); i++){
            newMatrix.setRow(i, (List) matrix.get(i));
        }
        return newMatrix;
    }

}
