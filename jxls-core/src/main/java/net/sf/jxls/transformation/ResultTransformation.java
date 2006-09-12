package net.sf.jxls.transformation;

import net.sf.jxls.transformation.BaseTransformation;
import net.sf.jxls.tag.Block;

/**
 * Result information about transformation
 * @author Leonid Vysochyn
 */
public class ResultTransformation extends BaseTransformation {
    int lastRowShift;
    int nextRowShift;
    short nextCellShift;
    short lastCellShift;

    public ResultTransformation() {
    }

    public ResultTransformation(short nextCellShift, short lastCellShift) {
        this.nextCellShift = nextCellShift;
        this.lastCellShift = lastCellShift;
    }

    public ResultTransformation(int nextRowShift) {
        this.nextRowShift = nextRowShift;
    }

    public ResultTransformation(int nextRowShift, int lastRowShift) {
        this.nextRowShift = nextRowShift;
        this.lastRowShift = lastRowShift;
    }

    public ResultTransformation add(ResultTransformation transformation){
        lastRowShift += transformation.getLastRowShift();
        nextRowShift += transformation.getNextRowShift();
        lastCellShift += transformation.getLastCellShift();
        nextCellShift += transformation.getNextCellShift();
        return this;
    }

    public Block transformBlock(Block block){
        if( block!=null ){
            block = block.horizontalShift( lastCellShift );
            block = block.verticalShift( lastRowShift );
        }
        return block;
    }

    public ResultTransformation addNextRowShift( int shift ){
        nextRowShift += shift;
        return this;
    }

    public ResultTransformation addRightShift( short shift ){
        lastCellShift += shift;
        return this;
    }

    public short getLastCellShift() {
        return lastCellShift;
    }

    public int getLastRowShift() {
        return lastRowShift;
    }

    public void setLastRowShift(int lastRowShift) {
        this.lastRowShift = lastRowShift;
    }

    public int getNextRowShift() {
        return nextRowShift;
    }

    public void setNextRowShift(int nextRowShift) {
        this.nextRowShift = nextRowShift;
    }

    public short getNextCellShift() {
        return nextCellShift;
    }

}
