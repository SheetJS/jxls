package net.sf.jxls;

import junit.framework.TestCase;

import java.util.List;

import net.sf.jxls.controller.TransformationMatrix;
import net.sf.jxls.controller.MatrixCell;

/**
 * @author Leonid Vysochyn
 */
public class TransformationMatrixTest extends TestCase {
    public void testDuplicateRows(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.duplicateRows( 5, 7, 3 );
        for(int i = 8; i <= 10; i++){
            for(int j = 0; j < matrix.getColCount(); j++){
                assertEquals( "Duplicated row items must be equal. Items (" + (i - 3) + "," + j + ") and (" + i + "," + j + ") are not equal", origMatrix.get(i-3, j), matrix.get(i, j) );
            }
        }
        for(int i = 11; i <= 13; i++){
            for(int j = 0; j < matrix.getColCount(); j++){
                assertEquals( "Duplicated row items must be equal. Items (" + (i - 6) + "," + j + ") and (" + i + "," + j + ") are not equal", origMatrix.get(i-6, j), matrix.get(i, j) );
            }
        }
        for(int i = 14; i <= 16; i++){
            for(int j = 0; j < matrix.getColCount(); j++){
                assertEquals( "Duplicated row items must be equal. Items (" + (i - 9) + "," + j + ") and (" + i + "," + j + ") are not equal", origMatrix.get(i-9, j), matrix.get(i, j) );
            }
        }
        for(int i = 0; i <= 7; i++){
            for(int j = 0; j < matrix.getColCount(); j++){
                assertEquals("Not duplicated rows must have been left untouched", origMatrix.get(i,j), matrix.get(i,j) );
            }
        }
        for(int i = 8; i <= 19; i++){
            for(int j = 0; j < matrix.getColCount(); j++){
                assertEquals("Not duplicated rows must have been left untouched", origMatrix.get(i,j), matrix.get(i + 9, j) );
            }
        }
    }

    public void testShiftRowsDown(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.shiftRows( 5, 3 );
        assertEquals( "Number of matrix rows after shift is incorrect", origMatrix.getRowCount() + 3, matrix.getRowCount() );
        for(int i = 5; i < 20; i++){
            for(int j = 0; j < origMatrix.getColCount(); j++){
                assertEquals( "Shifted row items must be equal. Items (" + i + "," + j + ") and (" + (i + 3) + "," + j + ") are not equal", origMatrix.get(i, j), matrix.get(i+3, j) );
            }
        }
    }
    public void testShiftRowsUp(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.shiftRows( 5, -3 );
        assertEquals( "Number of matrix rows after shift is incorrect", origMatrix.getRowCount() - 3, matrix.getRowCount() );
        for(int i = 5; i < 20; i++){
            for(int j = 0; j < origMatrix.getColCount(); j++){
                assertEquals( "Shifted row items must be equal. Items (" + i + "," + j + ") and (" + (i + 3) + "," + j + ") are not equal", origMatrix.get(i, j), matrix.get(i-3, j) );
            }
        }
    }

    public void testShiftDown(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.shift( 5, 10, 15, 3 );
        assertEquals( "Number of matrix rows after shift is incorrect", origMatrix.getRowCount() + 3, matrix.getRowCount() );
        for(int i = 5; i < origMatrix.getRowCount(); i++){
            for(int j = 10; j <= 15; j++){
                assertEquals( "Shifted row range items must be equal. Items (" + i + "," + j + ") and (" + (i + 3) + "," + j + ") are not equal", origMatrix.get(i, j), matrix.get(i+3, j) );
            }
        }
        checkMatrixRange( origMatrix, 0, origMatrix.getRowCount() - 1, 0, 9, matrix, 0, 0);
        checkMatrixRange( origMatrix, 0, origMatrix.getRowCount() - 1, 16, origMatrix.getColCount() - 1, matrix, 0, 16 );
    }

    public void testFindMappedCellsAfterShiftDown(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.shift( 5, 5, 10, 3 );
        List mappedCells = matrix.findMappedCells( 6, 7 );
        assertEquals("Number of mapped cells is incorrect", 1, mappedCells.size() );
        assertEquals("Mapped cell is incorrect", new MatrixCell(9, 7), mappedCells.get(0));

    }

    public void testShiftUp(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.shift( 5, 10, 15, -3 );
        assertEquals( "Number of matrix rows after shift is incorrect", origMatrix.getRowCount(), matrix.getRowCount() );
        for(int i = 5; i < origMatrix.getRowCount(); i++){
            for(int j = 10; j <= 15; j++){
                assertEquals( "Shifted row range items must be equal. Items (" + i + "," + j + ") and (" + (i + 3) + "," + j + ") are not equal", origMatrix.get(i, j), matrix.get(i - 3, j) );
            }
        }
        checkMatrixRange( origMatrix, 0, origMatrix.getRowCount() - 1, 0, 9, matrix, 0, 0);
        checkMatrixRange( origMatrix, 0, origMatrix.getRowCount() - 1, 16, origMatrix.getColCount() - 1, matrix, 0, 16 );
    }

    private void checkMatrixRange(TransformationMatrix srcMatrix, int srcStartRowNum, int srcEndRowNum, int srcStartColNum, int srcEndColNum,
                                  TransformationMatrix destMatrix, int destStartRowNum, int destStartColNum){
        int ii = destStartRowNum;
        int jj;
        for(int i = srcStartRowNum; i <= srcEndRowNum; i++){
            jj = destStartColNum;
            for(int j = srcStartColNum; j <= srcEndColNum; j++){
                assertEquals("Items from source matrix are not equal to items of destination matrix", srcMatrix.get( i, j ), destMatrix.get(ii, jj) );
                jj++;
            }
            ii++;
        }
    }

    public void testFindMappedCells(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        matrix.shiftRows( 5, 3 );
        checkSingleMappedCell( matrix, new MatrixCell(10, 1), new MatrixCell(13, 1) );
        matrix.shiftRows( 7, 5 );
        checkSingleMappedCell( matrix, new MatrixCell(10, 1), new MatrixCell(18, 1) );
        matrix.shiftRows( 10, -2 );
        checkSingleMappedCell( matrix, new MatrixCell(10, 1), new MatrixCell(16, 1) );
        matrix.duplicateRows( 10, 12, 2 );
        checkSingleMappedCell( matrix, new MatrixCell(2, 3), new MatrixCell(2, 3) );
        matrix.shift( 10, 5, 10, -2 );
        List mappedCells = matrix.findMappedCells( 6, 1 );
        assertEquals("Number of mapped cells is incorrect", 3, mappedCells.size() );
        assertTrue("Can't find mapped cell", mappedCells.indexOf( new MatrixCell(12, 1) ) >= 0);
        assertTrue("Can't find mapped cell", mappedCells.indexOf( new MatrixCell(15, 1) ) >= 0);
        assertTrue("Can't find mapped cell", mappedCells.indexOf( new MatrixCell(18, 1) ) >= 0);
        mappedCells = matrix.findMappedCells( 6, 7 );
        assertEquals("Number of mapped cells is incorrect", 3, mappedCells.size() );
        assertTrue("Can't find mapped cell", mappedCells.indexOf( new MatrixCell(10, 7) ) >= 0);
        assertTrue("Can't find mapped cell", mappedCells.indexOf( new MatrixCell(13, 7) ) >= 0);
        assertTrue("Can't find mapped cell", mappedCells.indexOf( new MatrixCell(16, 7) ) >= 0);
    }

    private void checkSingleMappedCell(TransformationMatrix matrix, MatrixCell srcCell, MatrixCell destCell){
        List mappedCells = matrix.findMappedCells( srcCell.getRowNum(), srcCell.getColNum() );
        assertEquals("Several mapped cells for source cell are not expected", 1, mappedCells.size() );
        MatrixCell cell = (MatrixCell) mappedCells.get(0);
        assertEquals("Mapped cell is incorrect", destCell, cell);
    }

    public void testDuplicateRight(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.duplicateRight( 5, 7, 10, 12, 3 );
        for(int i = 5; i <= 7; i++){
            for(int j = 10; j <= 12; j++){
                for(int k = 0; k <= 3; k++){
                    assertEquals( "Duplicated row items must be equal. Items (" + i + "," + j + ") and (" + i + "," + (j + 3*k) + ") are not equal", origMatrix.get(i, j), matrix.get(i, j+3*k) );
                }
            }
        }
    }

    public void testShiftColumns(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.shiftColumns( 5, 7, 10, 3 );
        for(int i = 5; i <= 7; i++){
            for(int j = 10; j < origMatrix.getColCount(); j++){
                assertEquals( "Right shifted items must be equal to original ones. Items (" + i + "," + j + ") and (" + i + "," + (j+3) + ") are not equal", origMatrix.get(i,j), matrix.get(i, j + 3) );
            }
        }
        matrix = (TransformationMatrix) origMatrix.clone();
        matrix.shiftColumns( 5, 7, 10, -3);
        for(int i = 5; i <= 7; i++){
            for(int j = 10; j < origMatrix.getColCount(); j++){
                assertEquals( "Right shifted items must be equal to original ones. Items (" + i + "," + j + ") and (" + i + "," + (j-3) + ") are not equal", origMatrix.get(i,j), matrix.get(i, j - 3) );
            }
        }
    }

    public void testDuplicate(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.duplicate( 5, 10, 10, 15, 5, false);
        for(int k = 0; k < 5; k++){
            int d = 6 + k * 6;
            for(int i = 5; i <= 10; i++){
                for(int j = 10; j <= 15; j++){
                    assertEquals( "Duplicated row items must be equal. Items (" + i + "," + j + ") and (" + (i + d) +  "," + j + ") are not equal", origMatrix.get(i, j), matrix.get(i + d, j) );
                }
            }
            for(int i = d + 5; i <= d + 10 && i < origMatrix.getRowCount(); i++){
                for(int j = 0; j <= 9; j++){
                    assertEquals( "Not duplicated row items must be equal. Item (" + i + "," + j + ") has changed ", origMatrix.get(i, j), matrix.get(i, j) );
                }
            }
            for(int i = d + 5; i <= d + 10 && i < origMatrix.getRowCount(); i++){
                for(int j = 16; j < matrix.getColCount(); j++){
                    assertEquals( "Not duplicated row items must be equal. Item (" + i + "," + j + ") has changed ", origMatrix.get(i, j), matrix.get(i, j) );
                }
            }
        }
    }

    public void testRemoveBorders(){
        TransformationMatrix matrix = new TransformationMatrix(20, 20);
        TransformationMatrix origMatrix = (TransformationMatrix) matrix.clone();
        matrix.removeBorders(5, 10, 10, 15);
        checkMatrixRange( origMatrix, 0, 4, 0, origMatrix.getColCount() - 1, matrix, 0, 0 );
        checkMatrixRange( origMatrix, 0, origMatrix.getRowCount() - 1, 0, 9, matrix, 0, 0 );
        checkMatrixRange( origMatrix, 6, 9, 11, 14, matrix, 5, 10 );
        //todo
//        checkMatrixRange( origMatrix, 5, 10, 11, 14, matrix, 5, 10 );

    }
}
